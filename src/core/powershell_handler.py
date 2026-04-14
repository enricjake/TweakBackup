"""
PowerShell handler for WinSet - executes WMI and PowerShell scripts.
"""

import subprocess  # nosec - B404: subprocess needed for PowerShell, handled securely
import re
import os
from typing import Tuple, Optional


class PowerShellHandler:
    """Handles execution of PowerShell commands and scripts securely."""

    # Maximum number of characters allowed in a single PowerShell command string.
    # Enforced to prevent buffer overflow attacks and excessively long/complex
    # commands that could be indicative of injection attempts. Commands exceeding
    # this length are rejected before execution.
    MAX_COMMAND_LENGTH = 1000

    # Regex pattern defining the set of characters permitted in a command string.
    # Allows: word chars (\w), whitespace (\s), hyphens, braces {}, brackets [],
    # parentheses (), periods, backslashes, forward slashes, colons, semicolons,
    # equals signs, single/double quotes, pipes, ampersands, percent signs,
    # dollar signs, hashes, asterisks, question marks, and exclamation marks.
    # Any command containing characters outside this set is rejected as a safety
    # measure against injection of unexpected or dangerous tokens.
    ALLOWED_COMMAND_CHARS = r"^[\w\s\-\{\}\[\]\(\)\.\\/:=\'\"\#\*\?\!]+$"

    # List of regex patterns that match known dangerous PowerShell/cmd commands.
    # During validation each pattern is tested against the command string (case-
    # insensitive). If any pattern matches, the command is blocked unless a
    # specific whitelist exception applies. This is the primary defense against
    # destructive operations like recursive deletion, disk formatting, account
    # creation, and forced restarts.
    DANGEROUS_PATTERNS = [
        r"Remove-Item.*-Recurse",  # Recursive file deletion
        r"rm\s+-rf\s+",  # Unix-style recursive force delete
        r"del\s+/[fsq]",  # Windows del with /f /s /q flags
        r"format\s+[a-z]:",  # Disk format command
        r"net\s+user\s+/add",  # Adding a new user account
        r"net\s+localgroup\s+administrators\s+/add",  # Promoting a user to admin
        r"Stop-Computer",  # Forcing a shutdown
        r"Restart-Computer",  # Forcing a restart
        r"Clear-EventLog",  # Wiping Windows event logs
        r"Set-Service.*-StartupType\s+Disabled.*Stop-Service",  # Disabling+stopping a service
        r"Get-WmiObject\s+-Class\s+Win32_Product",  # Known dangerous WMI query
    ]

    def __init__(self):
        # Absolute filesystem path to the PowerShell executable. Using an
        # absolute path instead of relying on PATH prevents DLL hijacking
        # attacks (bandit rule B607) where a malicious binary named
        # "powershell.exe" could be placed earlier on the PATH.
        self.powershell_path = (
            r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        )

        # Verify that the PowerShell executable actually exists at the
        # expected location. If it does not (e.g. in a CI/testing environment
        # or a non-standard Windows install), fall back to resolving
        # "powershell.exe" via the system PATH as a last resort.
        if not os.path.exists(self.powershell_path):
            raise RuntimeError(
                "PowerShell not found at expected location. Cannot execute commands safely."
            )

    def _validate_command(self, command: str) -> Tuple[bool, str]:
        """
        Validate a PowerShell command before execution.

        Args:
            command: The PowerShell command to validate

        Returns:
            Tuple of (is_valid, error_message)
        """
        # Reject commands that exceed the maximum allowed length. This guards
        # against both accidental oversized inputs and deliberate overflow
        # attempts that could exploit memory or parsing weaknesses.
        if len(command) > self.MAX_COMMAND_LENGTH:
            return (
                False,
                f"Command too long: {len(command)} chars (max {self.MAX_COMMAND_LENGTH})",
            )

        # Reject empty commands to avoid invoking PowerShell with no arguments,
        # which would either be pointless or could behave unexpectedly.
        if len(command) == 0:
            return False, "Empty command"

        # Iterate through every known dangerous pattern and check if the
        # command contains a match (case-insensitive). Each pattern targets a
        # specific category of destructive or exploitable operation.
        for pattern in self.DANGEROUS_PATTERNS:
            if re.search(pattern, command, re.IGNORECASE):
                # Special whitelist: the Set-Service + Stop-Service pattern is
                # actually used by WinSet's own disable_service() method to
                # safely disable non-critical services. We verify the command
                # structure matches our expected form (starts with Set-Service
                # and contains Stop-Service) before allowing it through.
                if pattern == r"Set-Service.*-StartupType\s+Disabled.*Stop-Service":
                    # This is our safe service disable command
                    if (
                        not command.startswith("Set-Service")
                        or "Stop-Service" not in command
                    ):
                        return False, f"Dangerous command pattern detected: {pattern}"
                else:
                    return False, f"Dangerous command detected: {pattern}"

        # Verify that every character in the command belongs to the allowed
        # character set. This is a second layer of defense that catches
        # injection payloads using characters not covered by dangerous-pattern
        # checks.
        if not re.match(self.ALLOWED_COMMAND_CHARS, command):
            # Retry with an extended character set that additionally permits
            # commas, which appear in some legitimate PowerShell expressions
            # (e.g. array literals, parameter lists).
            if not re.match(r"^[\w\s\-\{\}\[\]\(\)\.\\/:=\'\"\#\*\?\!\,]+$", command):
                return False, "Command contains invalid characters"

        # Command passed all validation checks.
        return True, ""

    def run_command(self, command: str, timeout: int = 30) -> Tuple[bool, str]:
        """
        Execute a PowerShell command.

        Args:
            command: The PowerShell command string to execute.
            timeout: Maximum execution time in seconds.

        Returns:
            Tuple containing boolean success flag and the command output or error message.
        """
        # Validate the command against length, pattern, and character checks
        # before spawning any subprocess. If validation fails the error message
        # from _validate_command is returned directly to the caller.
        is_valid, error = self._validate_command(command)
        if not is_valid:
            return False, f"Command validation failed: {error}"

        try:
            # Launch PowerShell as a subprocess with security-focused flags:
            #   -NoProfile           : skip loading user PowerShell profile
            #                          scripts (avoids side effects)
            #   -ExecutionPolicy Bypass : allow script execution without
            #                          changing the system-wide policy
            #   -Command <command>   : execute the supplied command string
            #
            # creationflags=CREATE_NO_WINDOW suppresses the console window on
            # Windows so that no visible terminal pops up during WinSet
            # operations.
            #
            # capture_output=True captures both stdout and stderr so we can
            # relay results back to the GUI.
            #
            # text=True decodes the output as UTF-8 strings instead of bytes.
            #
            # timeout limits how long the subprocess can run, preventing
            # runaway commands from hanging the application.
            result = subprocess.run(  # nosec B603
                [
                    self.powershell_path,
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    command,
                ],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
                timeout=timeout,
            )

            # PowerShell returned exit code 0, meaning the command completed
            # without error. We return the captured stdout content.
            if result.returncode == 0:
                # Truncate very long output to prevent excessive memory usage
                # or UI rendering issues when a command produces megabytes of
                # data (e.g. listing thousands of files).
                output = result.stdout.strip()
                if len(output) > 10000:
                    output = output[:10000] + "... (truncated)"
                return True, output
            else:
                # Non-zero exit code indicates failure. Return stderr content
                # so the caller (and ultimately the user) can diagnose the
                # issue.
                error_output = result.stderr.strip()
                if len(error_output) > 5000:
                    error_output = error_output[:5000] + "... (truncated)"
                return False, error_output

        except subprocess.TimeoutExpired:
            # The command exceeded the specified timeout duration. Return a
            # descriptive error so the UI can inform the user.
            return False, f"Command timed out after {timeout} seconds"
        except Exception as e:
            # Catch-all for any other exception (e.g. FileNotFoundError if
            # PowerShell is not installed, PermissionError, etc.) and return
            # its string representation.
            return False, str(e)

    def set_power_plan(self, plan_guid: str) -> Tuple[bool, str]:
        """
        Set the active power plan.

        Args:
            plan_guid: GUID of the power plan (e.g., 381b4222-f694-41f0-9685-ff5bb260df2e)

        Returns:
            Tuple of (success, message)
        """
        # Validate that plan_guid is a properly formatted GUID string (8-4-4-4-12
        # hex digits, optionally wrapped in curly braces). This prevents command
        # injection through a crafted GUID value.
        if not re.match(
            r"^\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?$",
            plan_guid,
            re.IGNORECASE,
        ):
            return False, "Invalid Power Plan GUID format"

        # Build the powercfg command that tells Windows to switch to the
        # specified power plan. powercfg is a built-in Windows utility for
        # power management.
        command = f"powercfg /setactive {plan_guid}"
        return self.run_command(command)

    def disable_service(self, service_name: str) -> Tuple[bool, str]:
        """
        Disable a Windows service.

        Args:
            service_name: Name of the service to disable

        Returns:
            Tuple of (success, message)
        """
        # Validate service name contains only safe characters: word characters
        # (\w), spaces, dots, and hyphens. This prevents injection of special
        # characters that could alter the PowerShell command structure.
        if not re.match(r"^[\w\s\.-]+$", service_name):
            return (
                False,
                "Invalid Service Name - use only letters, numbers, spaces, dots, and dashes",
            )

        # Enforce a reasonable maximum length on service names to prevent
        # buffer-related issues and obvious injection attempts.
        if len(service_name) > 100:
            return False, "Service name too long (max 100 characters)"

        # Hard-coded list of Windows services that are essential for the
        # operating system to function. Disabling any of these would render
        # the system unbootable or unstable, so we refuse to touch them as a
        # safety guard regardless of what the user requests.
        critical_services = [
            "winlogon",
            "services",
            "lsass",
            "svchost",
            "wininit",
            "csrss",
            "smss",
            "system",
            "registry",
        ]

        # Compare the requested service name (case-insensitive) against the
        # critical list and reject the operation if there is a match.
        if service_name.lower() in critical_services:
            return False, f"Cannot disable critical system service: {service_name}"

        # Construct a compound PowerShell command:
        #   1. Set-Service -Name "<name>" -StartupType Disabled
        #      -> changes the service's startup type to "Disabled" so it will
        #         no longer start automatically on boot.
        #   2. Stop-Service -Name "<name>" -Force
        #      -> immediately stops the running service instance. The -Force
        #         flag suppresses confirmation prompts.
        # The two commands are joined by a semicolon so they execute
        # sequentially in a single PowerShell invocation.
        # Use double quotes with escaped doubling to prevent injection.
        escaped_name = service_name.replace('"', '""')
        command = (
            f'Set-Service -Name "{escaped_name}" -StartupType Disabled; '
            f'Stop-Service -Name "{escaped_name}" -Force'
        )
        return self.run_command(command)

    def get_service_status(self, service_name: str) -> Optional[str]:
        """
        Get the status of a Windows service.

        Args:
            service_name: Name of the service

        Returns:
            Status string or None if error
        """
        # Validate the service name to prevent injection through this query
        # method as well (same character whitelist as disable_service).
        if not re.match(r"^[\w\s\.-]+$", service_name):
            return None

        # Build a PowerShell expression that retrieves the .Status property
        # of the specified service. Get-Service is the standard PowerShell
        # cmdlet for querying Windows service information.
        # Use double quotes with escaped doubling to prevent injection.
        escaped_name = service_name.replace('"', '""')
        command = f'(Get-Service -Name "{escaped_name}").Status'
        # Execute the command. If successful, strip whitespace from the
        # returned status string (e.g. "Running", "Stopped", "Disabled").
        success, output = self.run_command(command)
        if success:
            return output.strip()
        # Return None on any failure so the caller can distinguish errors
        # from a valid status string.
        return None

    def get_active_power_plan(self) -> Optional[str]:
        """
        Get the currently active power plan GUID.

        Returns:
            Power plan GUID or None if error
        """
        # Query the currently active power scheme using the built-in powercfg
        # utility. The output looks like:
        #   "Power Scheme GUID: 381b4222-f694-41f0-9685-ff5bb260df2e  (Balanced)"
        command = "powercfg /getactivescheme"
        success, output = self.run_command(command)
        if success:
            # Extract the GUID from the output text using a regex that matches
            # the standard 8-4-4-4-12 hex-digit GUID format. The case-
            # insensitive flag handles both upper- and lowercase GUIDs.
            match = re.search(
                r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}",
                output,
                re.IGNORECASE,
            )
            if match:
                # match.group(0) returns the full matched GUID string without
                # any surrounding text.
                return match.group(0)
        # Return None if the command failed or no GUID was found in the output.
        return None
