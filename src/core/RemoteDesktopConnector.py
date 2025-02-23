import subprocess
import os
import platform
import time

class RemoteDesktopConnector:
    def __init__(self, ip_address):
        self.ip_address = ip_address
        self.is_wsl_env = self.is_wsl()

    def connect(self):
        """
        Detecta el entorno y usa el método apropiado para conectar
        """
        if self.is_wsl_env:
            return self.connect_from_wsl()
        else:
            return self.connect_from_windows()

    def disconnect(self):
        """
        Cierra la conexión RDP y limpia las credenciales
        """
        try:
            if self.is_wsl_env:
                # Cerrar todas las sesiones RDP que coincidan con la IP
                disconnect_command = ['powershell.exe', '-Command', f'''
                    $processes = Get-Process | Where-Object {{
                        $_.ProcessName -eq 'mstsc' -and (
                            $_.MainWindowTitle -match '{self.ip_address}' -or
                            $_.MainWindowTitle -match 'Remote Desktop Connection'
                        )
                    }}
                    if ($processes) {{
                        $processes | ForEach-Object {{ $_.Kill() }}
                    }}
                ''']
                subprocess.run(disconnect_command, capture_output=True)
                
                # Eliminar credenciales
                delete_cred_command = ['powershell.exe', '-Command',
                                     f'cmdkey.exe /delete:{self.ip_address}']
                subprocess.run(delete_cred_command, capture_output=True)
            else:
                # Versión para Windows nativo
                subprocess.run(f'taskkill /FI "WINDOWTITLE eq {self.ip_address}*" /IM mstsc.exe /F', 
                             capture_output=True)
                subprocess.run(f'cmdkey /delete:{self.ip_address}', capture_output=True)
            
            return True
        except Exception as e:
            print(f"Error al desconectar: {e}")
            return False

    def connect_from_windows(self):
        """
        Conecta usando el método directo de Windows
        """
        try:
            subprocess.run(f"cmdkey /generic:{self.ip_address} /user:LocalAccount\\instrument /pass:894129", 
                         capture_output=True)
            subprocess.Popen(f"mstsc /v:{self.ip_address} /f", 
                           close_fds=True, 
                           creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0)
            return True
        except Exception as e:
            print(f"Error al conectar desde Windows: {e}")
            return False

    def connect_from_wsl(self):
        """
        Conecta usando el método WSL
        """
        try:
            # Cambiamos al directorio temporal de Windows
            windows_temp = "/mnt/c/Windows/Temp"
            os.chdir(windows_temp)
            
            # Agregar credenciales silenciosamente
            cmdkey_command = ['powershell.exe', '-Command', 
                            f'cmdkey.exe /generic:{self.ip_address} /user:LocalAccount\\instrument /pass:894129']
            subprocess.run(cmdkey_command, capture_output=True)
            
            # Iniciar RDP
            mstsc_command = ['powershell.exe', '-Command', 
                           f'Start-Process mstsc.exe -ArgumentList "/v:{self.ip_address}","/f"']
            subprocess.run(mstsc_command, capture_output=True)
            
            return True
            
        except Exception as e:
            print(f"Error inesperado: {e}")
            return False

    @staticmethod
    def is_wsl():
        """
        Detecta si estamos en WSL usando múltiples métodos
        """
        try:
            with open('/proc/version', 'r') as f:
                if 'microsoft' in f.read().lower():
                    return True
        except FileNotFoundError:
            pass

        if 'WSL_DISTRO_NAME' in os.environ:
            return True

        if platform.system() == 'Linux':
            try:
                with open('/proc/sys/kernel/osrelease', 'r') as f:
                    if 'microsoft' in f.read().lower():
                        return True
            except FileNotFoundError:
                pass

        return False