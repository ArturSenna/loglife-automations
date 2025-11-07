"""
Version Manager for LogLife Automations
Handles version checking and update notifications
"""
import json
import os
import sys
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import requests

# Current version
CURRENT_VERSION = "1.0.0"
BUILD_DATE = "2025-11-07"

# Update configuration
UPDATE_CHECK_URL = "https://raw.githubusercontent.com/ArturSenna/loglife-automations/main/version.json"
RELEASE_URL = "https://github.com/ArturSenna/loglife-automations/releases/latest"

class VersionManager:
    def __init__(self):
        self.current_version = CURRENT_VERSION
        self.build_date = BUILD_DATE
        self.version_file = self._get_version_file_path()
        
    def _get_version_file_path(self):
        """Get the path to version file"""
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            base_path = Path(sys._MEIPASS)
        else:
            # Running as script
            base_path = Path(__file__).parent.parent
        
        return base_path / "version_info.json"
    
    def get_local_version_info(self):
        """Get local version information"""
        return {
            "version": self.current_version,
            "build_date": self.build_date,
            "last_check": datetime.now().isoformat()
        }
    
    def check_for_updates(self, show_no_update_message=False):
        """
        Check for updates from GitHub
        Returns: tuple (has_update, remote_version, download_url)
        """
        try:
            response = requests.get(UPDATE_CHECK_URL, timeout=10)
            response.raise_for_status()
            remote_info = response.json()
            
            remote_version = remote_info.get("version", "0.0.0")
            download_url = remote_info.get("download_url", RELEASE_URL)
            update_notes = remote_info.get("notes", "")
            
            has_update = self._compare_versions(remote_version, self.current_version)
            
            if has_update:
                return True, remote_version, download_url, update_notes
            elif show_no_update_message:
                messagebox.showinfo(
                    "Atualização",
                    f"Você está usando a versão mais recente ({self.current_version})"
                )
            
            return False, remote_version, download_url, update_notes
            
        except requests.RequestException as e:
            if show_no_update_message:
                messagebox.showwarning(
                    "Verificação de Atualização",
                    f"Não foi possível verificar atualizações.\nErro: {str(e)}"
                )
            return False, None, None, None
        except Exception as e:
            if show_no_update_message:
                messagebox.showerror(
                    "Erro",
                    f"Erro ao verificar atualizações: {str(e)}"
                )
            return False, None, None, None
    
    def _compare_versions(self, remote, local):
        """
        Compare version strings
        Returns True if remote is newer than local
        """
        def version_tuple(v):
            return tuple(map(int, v.split('.')))
        
        try:
            return version_tuple(remote) > version_tuple(local)
        except:
            return False
    
    def prompt_update(self, remote_version, download_url, update_notes):
        """Show update dialog to user"""
        message = f"""Nova versão disponível!

Versão atual: {self.current_version}
Nova versão: {remote_version}

{update_notes}

Deseja baixar a atualização?"""
        
        result = messagebox.askyesno(
            "Atualização Disponível",
            message
        )
        
        if result:
            import webbrowser
            webbrowser.open(download_url)
            messagebox.showinfo(
                "Atualização",
                "O navegador será aberto para download.\n\n"
                "Após o download:\n"
                "1. Feche este programa\n"
                "2. Substitua o arquivo executável\n"
                "3. Execute a nova versão"
            )
    
    def force_update_check(self):
        """Force an update check with user feedback"""
        has_update, remote_version, download_url, update_notes = self.check_for_updates(show_no_update_message=True)
        
        if has_update:
            self.prompt_update(remote_version, download_url, update_notes)
    
    def auto_check_on_startup(self):
        """Automatically check for updates on startup (silent unless update found)"""
        has_update, remote_version, download_url, update_notes = self.check_for_updates(show_no_update_message=False)
        
        if has_update:
            self.prompt_update(remote_version, download_url, update_notes)


def get_version_info():
    """Get current version information as string"""
    return f"Versão {CURRENT_VERSION} (Build: {BUILD_DATE})"


if __name__ == "__main__":
    # Test the version manager
    vm = VersionManager()
    print(get_version_info())
    vm.force_update_check()
