# ctypes: Provides access to the Windows API, used here to check admin privileges
# via the Windows shell32 DLL (IsUserAnAdmin).
import ctypes

# subprocess: Used to execute external PowerShell commands to create System Restore
# points. Marked with # nosec to suppress Bandit security warnings, since the
# command is hardened (absolute path, sanitized input).
import subprocess  # nosec

# os: Standard library module for OS-level operations (not actively used but
# imported for potential future path or environment handling).
import base64


class BackupManager:
    """Manages Windows System Restore points for safe application rollbacks."""

    @staticmethod
    def is_admin() -> bool:
        """Check if the script is running with administrator privileges.

        Returns:
            True if the current process has Administrator privileges,
            False otherwise (including if the Windows API call fails).
        """
        try:
            # Use IsElevated() which is more reliable than IsUserAnAdmin
            # IsElevated returns a non-zero value if the current process
            # is running with elevated privileges (UAC admin or full admin).
            return ctypes.windll.shell32.IsElevated() != 0
        except Exception:
            # Fallback to IsUserAnAdmin if IsElevated is not available
            try:
                return ctypes.windll.shell32.IsUserAnAdmin() != 0
            except Exception:
                return False

    def create_restore_point(
        self, description: str = "WinSet Configuration Backup"
    ) -> bool:
        """
        Creates a Windows System Restore point.
        Requires Administrator privileges.

        Args:
            description: A human-readable label for the restore point that
                         will appear in Windows System Restore listings.
                         Defaults to "WinSet Configuration Backup".

        Returns:
            True if the restore point was created successfully, False otherwise.
        """
        # --- Privilege gate ---
        # System Restore point creation requires elevated privileges. If the
        # current process is not running as Administrator, abort early with a
        # user-friendly message.
        if not self.is_admin():
            print("Cannot create restore point: Administrator privileges required.")
            return False

        try:
            # --- Input sanitization ---
            # Double any double-quotes in the user-supplied description so that
            # when interpolated into the PowerShell string literal below, the
            # value cannot break out of the quoted context and inject arbitrary
            # commands (PowerShell injection prevention).
            sanitized_desc = description.replace('"', '""')

            # --- PowerShell command construction ---
            # Checkpoint-Computer is the PowerShell cmdlet that creates a
            # System Restore point.
            #   -Description: the label shown in System Restore UI.
            #   -RestorePointType "MODIFY_SETTINGS": marks the restore point as
            #     a configuration-change type (appropriate for WinSet tweaks).
            #   -ErrorAction Stop: ensures non-terminating errors are raised as
            #     terminating errors so we can detect failures via exit code.
            ps_script = (
                f'Checkpoint-Computer -Description "{sanitized_desc}" '
                f'-RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop'
            )

            # --- PowerShell path ---
            # Hardcoded absolute path to the PowerShell executable. Using a
            # hardcoded path prevents PATH hijacking attacks (B607) where a
            # malicious executable named "powershell" could be placed earlier
            # on the system PATH and executed instead.
            # nosec B607: suppressed because we use an absolute path.
            powershell_path = (
                r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
            )

            # --- Subprocess execution ---
            # Arguments passed as a list (no shell=True) to avoid shell injection.
            #   -NoProfile: skips loading the PowerShell profile for faster,
            #     cleaner execution without user customizations.
            #   -ExecutionPolicy Bypass: ensures the command runs even if the
            #     system execution policy would otherwise block script execution.
            #   -Command: tells PowerShell the next argument is a command string.
            # capture_output=True: captures stdout and stderr for logging.
            # text=True: returns output as str instead of bytes.
            # CREATE_NO_WINDOW: prevents a visible console window from popping
            #   up on the user's desktop (Windows-only flag).
            # nosec B603: suppressed because the command array uses a hardcoded
            #   path and sanitized arguments — no shell=True.
            encoded = base64.b64encode(ps_script.encode("utf-16-le")).decode("ascii")
            result = subprocess.run(  # nosec B603
                [
                    powershell_path,
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-EncodedCommand",
                    encoded,
                ],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )

            # --- Result interpretation ---
            # returncode == 0 means the PowerShell cmdlet succeeded.
            if result.returncode == 0:
                print(f"Successfully created system restore point: {description}")
                return True
            else:
                # Non-zero exit code indicates failure; stderr contains the
                # error details from PowerShell.
                print(
                    f"Failed to create restore point. Details: {result.stderr.strip()}"
                )
                return False

        except Exception as e:
            # Catch-all for any unexpected Python-level exception (e.g.,
            # FileNotFoundError if PowerShell is not installed, PermissionError,
            # etc.). The exception message 'e' is printed for debugging.
            print(f"Exception during restore point creation: {e}")
            return False
