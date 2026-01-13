"""
SlideSnap Installer
Creates desktop shortcut, Start Menu entry, and registry uninstall entry
"""

import os
import sys
import shutil
import winreg
import subprocess
import tkinter as tk
from tkinter import messagebox
import threading

APP_NAME = "SlideSnap"
PUBLISHER = "SlideSnap"
VERSION = "1.0.0"


def get_resource_path(filename):
    """Get path to bundled resource."""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)


def create_shortcut(target_path, shortcut_path, icon_path=None, description=""):
    """Create Windows shortcut using PowerShell."""
    ps_script = f'''
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("{shortcut_path}")
$Shortcut.TargetPath = "{target_path}"
$Shortcut.WorkingDirectory = "{os.path.dirname(target_path)}"
$Shortcut.Description = "{description}"
'''
    if icon_path and os.path.exists(icon_path):
        ps_script += f'$Shortcut.IconLocation = "{icon_path}"\n'
    ps_script += '$Shortcut.Save()'

    subprocess.run(
        ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
        capture_output=True,
        creationflags=subprocess.CREATE_NO_WINDOW
    )


class InstallerWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("SlideSnap Installer")
        self.root.geometry("400x150")
        self.root.resizable(False, False)
        self.root.configure(bg='#1a1025')

        # Center on screen
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - 400) // 2
        y = (self.root.winfo_screenheight() - 150) // 2
        self.root.geometry(f"400x150+{x}+{y}")

        # Try to set icon
        try:
            icon_path = get_resource_path("slidesnap.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except:
            pass

        # Title
        title = tk.Label(
            self.root,
            text="Installing SlideSnap...",
            font=("Segoe UI", 14, "bold"),
            fg='#ff6ac1',
            bg='#1a1025'
        )
        title.pack(pady=(20, 10))

        # Status label
        self.status_label = tk.Label(
            self.root,
            text="Preparing installation...",
            font=("Segoe UI", 10),
            fg='#a89bb5',
            bg='#1a1025'
        )
        self.status_label.pack(pady=(0, 15))

        # Progress bar frame
        progress_frame = tk.Frame(self.root, bg='#3d2a54', height=8)
        progress_frame.pack(fill='x', padx=40)
        progress_frame.pack_propagate(False)

        self.progress_bar = tk.Frame(progress_frame, bg='#AF29F5', height=8, width=0)
        self.progress_bar.pack(side='left', fill='y')

        self.progress_width = 320  # Total width
        self.install_dir = None
        self.exe_path = None

    def update_status(self, text, progress=0):
        """Update status text and progress bar."""
        self.status_label.config(text=text)
        new_width = int(self.progress_width * progress / 100)
        self.progress_bar.config(width=new_width)
        self.root.update()

    def do_install(self):
        """Run installation in background."""
        try:
            # Install paths
            program_files = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
            self.install_dir = os.path.join(program_files, APP_NAME)

            self.update_status("Creating directories...", 10)
            os.makedirs(self.install_dir, exist_ok=True)

            # Copy app files
            self.update_status("Copying application files...", 20)
            app_source = get_resource_path("SlideSnap")
            if os.path.exists(app_source):
                items = os.listdir(app_source)
                for i, item in enumerate(items):
                    src = os.path.join(app_source, item)
                    dst = os.path.join(self.install_dir, item)
                    try:
                        if os.path.isdir(src):
                            if os.path.exists(dst):
                                shutil.rmtree(dst)
                            shutil.copytree(src, dst)
                        else:
                            shutil.copy2(src, dst)
                    except Exception:
                        pass
                    # Update progress (20-60%)
                    progress = 20 + int(40 * (i + 1) / len(items))
                    self.update_status(f"Copying files... ({i+1}/{len(items)})", progress)
            else:
                self.root.after(0, lambda: self.show_error("App source not found"))
                return

            self.exe_path = os.path.join(self.install_dir, "SlideSnap.exe")
            icon_path = os.path.join(self.install_dir, "slidesnap.ico")

            # Create desktop shortcut
            self.update_status("Creating desktop shortcut...", 65)
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            desktop_shortcut = os.path.join(desktop, f"{APP_NAME}.lnk")
            create_shortcut(self.exe_path, desktop_shortcut, icon_path, "Extract slides from videos")

            # Create Start Menu shortcut
            self.update_status("Creating Start Menu entry...", 75)
            start_menu = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs")
            start_menu_shortcut = os.path.join(start_menu, f"{APP_NAME}.lnk")
            create_shortcut(self.exe_path, start_menu_shortcut, icon_path, "Extract slides from videos")

            # Create uninstall script
            self.update_status("Creating uninstaller...", 85)
            uninstall_script = os.path.join(self.install_dir, "uninstall.bat")
            with open(uninstall_script, 'w') as f:
                f.write('@echo off\n')
                f.write('echo Uninstalling SlideSnap...\n')
                f.write('taskkill /F /IM SlideSnap.exe 2>nul\n')
                f.write('timeout /t 2 /nobreak >nul\n')
                f.write(f'del "{desktop_shortcut}" 2>nul\n')
                f.write(f'del "{start_menu_shortcut}" 2>nul\n')
                f.write('reg delete "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\SlideSnap" /f 2>nul\n')
                f.write(f'cd /d "%TEMP%"\n')
                f.write(f'rmdir /s /q "{self.install_dir}" 2>nul\n')
                f.write('echo SlideSnap has been uninstalled.\n')
                f.write('pause\n')

            # Add to Windows uninstall registry
            self.update_status("Registering with Windows...", 95)
            try:
                reg_path = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\SlideSnap"
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_path) as key:
                    winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, APP_NAME)
                    winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, VERSION)
                    winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, PUBLISHER)
                    winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, self.install_dir)
                    winreg.SetValueEx(key, "UninstallString", 0, winreg.REG_SZ, f'cmd.exe /c "{uninstall_script}"')
                    winreg.SetValueEx(key, "DisplayIcon", 0, winreg.REG_SZ, icon_path)
                    winreg.SetValueEx(key, "NoModify", 0, winreg.REG_DWORD, 1)
                    winreg.SetValueEx(key, "NoRepair", 0, winreg.REG_DWORD, 1)
            except Exception:
                pass

            self.update_status("Installation complete!", 100)
            self.root.after(500, self.show_success)

        except Exception as e:
            self.root.after(0, lambda: self.show_error(str(e)))

    def show_success(self):
        """Show success dialog."""
        self.root.withdraw()
        launch = messagebox.askyesno(
            "Installation Complete",
            f"SlideSnap has been installed successfully!\n\n"
            f"Location: {self.install_dir}\n\n"
            f"You can launch it from:\n"
            f"  • Desktop shortcut\n"
            f"  • Start Menu\n\n"
            f"Launch SlideSnap now?"
        )
        if launch and self.exe_path:
            subprocess.Popen([self.exe_path], creationflags=subprocess.CREATE_NO_WINDOW)
        self.root.destroy()

    def show_error(self, message):
        """Show error dialog."""
        self.root.withdraw()
        messagebox.showerror("Installation Error", f"Installation failed: {message}")
        self.root.destroy()

    def run(self):
        """Start installation."""
        # Start installation in background after window shows
        self.root.after(100, self.do_install)
        self.root.mainloop()


if __name__ == "__main__":
    try:
        installer = InstallerWindow()
        installer.run()
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Installation Error", f"Installation failed: {e}")
        root.destroy()
