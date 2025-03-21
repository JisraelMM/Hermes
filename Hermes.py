import os
import subprocess
import zipfile
import urllib.request
import sys
import importlib
import tarfile
import warnings
import time
import re
import ctypes
import winreg
import platform

# console colors
errorStyle = "\033[91m"
warningStyle = "\033[93m"
normalStyle = "\033[0m"
titleStyle = "\033[34;1;3m"
promptStyle = "\033[96m"
successStyle = "\033[92m"
runStyle = "\x1B[38;2;255;128;0m"
citestyle = "\x1B[38;2;17;245;120m"
exitStyle = "\033[34;1;3m"



# installing packages
def install_and_import(package_name):
    print("Updating and installing some libraries needed for this program, please be patient...")
    try:
        importlib.import_module(package_name)
        print(successStyle + f"{package_name} is already installed ✔" + normalStyle)
    except ImportError:
        #print(f"{package_name} is not installed. Installing now...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', package_name])
        #print(successStyle + f"{package_name} installed successfully ✔" + normalStyle)

def verify_and_install_packages():
    required_packages = ['requests', 'tqdm']
    for package in required_packages:
        install_and_import(package)

# verifing if there are all the neccesary pacakges
verify_and_install_packages()

# import packages
import requests
from tqdm import tqdm

############################################################################################################################################


def get_latest_blast_version():
    try:
        print(titleStyle + "\n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        print("For more information see: \nhttps://blast.ncbi.nlm.nih.gov/doc/blast-help/downloadblastdata.html")
        print("\n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        print(warningStyle + "Establishing connection with NCBI page")
        print(successStyle + "\nDownloading BLAST+, be patient and make sure you have a stable internet connection...\n" + promptStyle )
        url = "https://ftp.ncbi.nlm.nih.gov/blast/executables/blast+/LATEST/"
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url, headers=headers, timeout=10).text

        match = re.search(r'ncbi-blast-(\d+\.\d+\.\d+)\+', response)
        if match:
            version = match.group(1)
            print(warningStyle + f"Última versión encontrada: {version}")
            return version
        else:
            print("No se pudo encontrar la versión más reciente de BLAST+. Se utilizará la versión predeterminada.")
    except requests.exceptions.RequestException as e:
        print(f"Error al obtener la versión más reciente de BLAST+: {e}")
    
    return "2.16.0"  # if the any error installling the verstion 2.16

def download_and_extract(url, destination):
    try:
        os.makedirs(destination, exist_ok=True)
        print(f"Creando el directorio de destino: {destination}")

        file_name = os.path.join(destination, url.split('/')[-1])
        headers = {'User-Agent': 'Mozilla/5.0'}

        print(f"Descargando BLAST+ desde: {url}")
        with requests.get(url, headers=headers, stream=True, timeout=20) as response:
            response.raise_for_status()
            total_size = int(response.headers.get('Content-Length', 0))
            block_size = 8192  # Se aumenta a 8 KB para mejorar la velocidad

            with open(file_name, 'wb') as f, tqdm(
                total=total_size, unit='B', unit_scale=True, desc='Descargando BLAST+'
            ) as tqdm_bar:
                for chunk in response.iter_content(chunk_size=block_size):
                    f.write(chunk)
                    tqdm_bar.update(len(chunk))

        print(successStyle + f"\nBLAST+ descargado correctamente en: {destination}")

        # unzziping
        print("Descomprimiendo archivos...")
        if tarfile.is_tarfile(file_name):
            with tarfile.open(file_name, "r:gz") as tar:
                safe_extract(tar, destination)
            print(f"Extracción completada en: {destination}")
        else:
            print("Error: El archivo descargado no es un .tar.gz válido.")
            return None

    except Exception as e:
        print(f"Error durante la descarga o descompresión: {e}")
        return None

    return file_name

def safe_extract(tar, path="."):
    """ Extrae archivos de forma segura para evitar vulnerabilidades. """
    abs_path = os.path.abspath(path)
    for member in tar.getmembers():
        member_path = os.path.abspath(os.path.join(path, member.name))
        if not member_path.startswith(abs_path):
            raise Exception("Advertencia: Se ha detectado una posible extracción insegura.")
    tar.extractall(path, filter="data")


def find_bin_path(root_directory):
    """ Busca la carpeta 'bin' dentro del directorio de instalación. """
    print("Buscando la carpeta 'bin' en el directorio de instalación...")
    for root, dirs, files in os.walk(root_directory):
        if 'bin' in dirs:
            bin_path = os.path.join(root, 'bin')
            print(f"\nCarpeta 'bin' encontrada en: {bin_path}")
            return bin_path
    print("Error: No se encontró la carpeta 'bin'.")
    return None

def add_blast_to_path_permanently(blast_path):
    """ Añade la carpeta 'bin' de BLAST+ al PATH del sistema en Windows. """
    try:
        print("Añadiendo la ruta de BLAST+ al PATH del sistema...")
        result = subprocess.run(
            ['powershell', '[System.Environment]::GetEnvironmentVariable("Path", "User")'],
            capture_output=True, text=True
        )
        user_path = result.stdout.strip()

        if blast_path not in user_path:
            new_path = user_path + ";" + blast_path
            subprocess.run(['powershell', f'[System.Environment]::SetEnvironmentVariable("Path", "{new_path}", "User")'])
            print(f"\nRuta {blast_path} añadida al PATH del sistema correctamente.")
        else:
            print("La ruta de BLAST ya está en el PATH del usuario.")
    except Exception as e:
        print(f"Error al modificar el PATH del sistema: {e}")

def download_blast():
    """ Descarga y configura BLAST+ en el sistema. """
    latest_version = get_latest_blast_version()
    print(f"Versión de BLAST+ a descargar: {latest_version}")

    url = f"https://ftp.ncbi.nlm.nih.gov/blast/executables/blast+/LATEST/ncbi-blast-{latest_version}+-x64-win64.tar.gz"
    destination = os.path.join(os.path.expanduser("~"), "blast")
    print(f"Se procederá a descargar desde: {url}")

    tar_file = download_and_extract(url, destination)

    if tar_file:
        bin_path = find_bin_path(destination)
        if bin_path:
            add_blast_to_path_permanently(bin_path)





###############################################################################################################################


def check_execution_policy():
    try:
        print("=====================================")
        print("       Starting the process")
        print("=====================================")

        # Verificar si el sistema es Windows 10
        is_windows10 = platform.system() == "Windows" and platform.release() == "10"

        if not is_windows10:
            # Intentar importar el módulo solo si NO es Windows 10
            load_module = subprocess.run(['powershell', 'Import-Module', 'Microsoft.PowerShell.Security'], capture_output=True, text=True)
            if load_module.returncode != 0:
                print(f"Error loading module Microsoft.PowerShell.Security: {load_module.stderr}")
                return False

        # Obtener la política de ejecución
        result = subprocess.run(['powershell', 'Get-ExecutionPolicy'], capture_output=True, text=True)
        policy = result.stdout.strip()

        if policy != 'Unrestricted':
            print(f"Current execution policy: {policy}")
            user_response = input("Would you like 'Hermes' to change the policy to 'Unrestricted'? (yes/no): ").strip().lower()

            if user_response == 'yes':
                print("Attempting to change the execution policy to 'Unrestricted'...")
                change_policy = subprocess.run(
                    ['powershell', 'Set-ExecutionPolicy', 'Unrestricted', '-Scope', 'CurrentUser', '-Force'],
                    capture_output=True, text=True
                )

                if change_policy.returncode == 0:
                    print("Execution policy successfully changed to Unrestricted ✔")
                    return True
                else:
                    print(f"Error changing execution policy: {change_policy.stderr}")
                    return False
            else:
                print("Execution policy will not be changed. Exiting the program...")
                return False
        else:
            print("\n✔ Execution policy is already set to 'Unrestricted'! ")
            return True
    except Exception as e:
        print(f"Error checking or changing execution policy: {e}")
        return False


###############################################################################################################################

# funtion to  Restricted
def change_execution_policy_to_restricted():
    try:
        print(warningStyle + "Attempting to change the execution policy to 'Restricted'..." + normalStyle)
        change_policy = subprocess.run(['powershell', 'Set-ExecutionPolicy', 'Restricted', '-Scope', 'CurrentUser', '-Force'], capture_output=True, text=True)
        
        if change_policy.returncode == 0:
            print(successStyle + "Execution policy successfully changed to Restricted ✔" + normalStyle)
        else:
            print(errorStyle + f"Error changing execution policy: {change_policy.stderr}" + normalStyle)
    except Exception as e:
        print(errorStyle + f"Error changing execution policy: {e}" + normalStyle)


###############################################################################################################################

# Función para actualizar pip
def upgrade_pip():
    try:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', 'pip'])
        print(successStyle + "pip upgraded successfully ✔" + normalStyle)
    except subprocess.CalledProcessError as e:
        print(errorStyle + f"Error upgrading pip: {e}" + normalStyle)

# Función para instalar paquetes
def install_package(package):
    try:
        print(f"Installing {package}...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
        print(successStyle + f"{package} installed successfully ✔" + normalStyle)
    except subprocess.CalledProcessError as e:
        print(errorStyle + f"Error installing {package}: {e}" + normalStyle)


##############################################################################################################################


def get_latest_vsearch_url():
    """Gets the URL of the latest VSEARCH version from GitHub."""
    try:
        api_url = "https://api.github.com/repos/torognes/vsearch/releases/latest"
        response = requests.get(api_url)
        response.raise_for_status()
        latest_release = response.json()
        for asset in latest_release["assets"]:
            if "win-x86_64.zip" in asset["browser_download_url"]:
                return asset["browser_download_url"]
        print(warningStyle + "No Windows binary found for the latest release." + normalStyle)
        return None
    except Exception as e:
        print(warningStyle + f"Error fetching latest VSEARCH release: {e}" + normalStyle)
        return None

def download_with_progress(url, destination):
    """Downloads a file with a progress bar."""
    def report_hook(block_num, block_size, total_size):
        downloaded = block_num * block_size
        percent = min(downloaded / total_size, 1.0)
        bar_length = 40
        block = int(bar_length * percent)
        progress_bar = "█" * block + "-" * (bar_length - block)
        sys.stdout.write(f"\rDownloading: [{progress_bar}] {percent * 100:.2f}%")
        sys.stdout.flush()
        
    urllib.request.urlretrieve(url, destination, reporthook=report_hook)
    print(successStyle + "\n✨ Download completed ✔")

def download_vsearch(url, extract_to):
    """Downloads and extracts VSEARCH in the specified folder."""
    try:
        os.makedirs(extract_to, exist_ok=True)
        zip_file = os.path.join(extract_to, "vsearch.zip")
        print(f"Downloading VSEARCH from {url}...")
        download_with_progress(url, zip_file)
        
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            print(runStyle + "✨ Extracting files...")
            zip_ref.extractall(extract_to)
        
        os.remove(zip_file)
        print(successStyle + "✨ Extraction completed successfully ✔" + normalStyle)
        
        for folder in os.listdir(extract_to):
            potential_path = os.path.join(extract_to, folder, "bin", "vsearch.exe")
            if os.path.exists(potential_path):
                return os.path.join(extract_to, folder, "bin")
        
        print(errorStyle + "vsearch.exe not found after extraction." + normalStyle)
        return None
    except Exception as e:
        print(errorStyle + f"Error downloading or extracting VSEARCH: {e}" + normalStyle)
        return None

def add_vsearch_to_path_permanently(vsearch_path):
    """Adds VSEARCH to the system PATH permanently."""
    try:
        result = subprocess.run(['powershell', '[System.Environment]::GetEnvironmentVariable("Path", "User")'], capture_output=True, text=True)
        user_path = result.stdout.strip()

        if vsearch_path not in user_path:
            new_path = user_path + ";" + vsearch_path
            subprocess.run(['powershell', f'[System.Environment]::SetEnvironmentVariable("Path", "{new_path}", "User")'])
            print(successStyle + "✨ VSEARCH has been successfully added to the system PATH! ✔" + normalStyle)
        else:
            print(successStyle + "✨ VSEARCH is already in the system PATH." + normalStyle)
    except Exception as e:
        print(errorStyle + f"Error adding VSEARCH to system PATH: {e}" + normalStyle)

def install_vsearch():
    clear_screen()
    print(promptStyle + "_______________________________________________________________________________________________")
    print("|                                                                                             |")
    print("|                     ***Running Hermes (v1.1.0)***                                           |")
    print("|_____________________________________________________________________________________________|")
    print("\n\n" + runStyle)

    """Main function to install and configure VSEARCH."""
    vsearch_url = get_latest_vsearch_url()
    if vsearch_url:
        destination = os.path.join(os.path.expanduser("~"), "VSEARCH")
        vsearch_path = download_vsearch(vsearch_url, destination)

        if vsearch_path:
            add_vsearch_to_path_permanently(vsearch_path)
            print(successStyle + "✨ VSEARCH installation completed successfully! ✔" + normalStyle)
        else:
            print(errorStyle + "vsearch.exe not found, cannot add to PATH." + normalStyle)
    else:
        print(errorStyle + "Failed to retrieve the latest VSEARCH download URL." + normalStyle)

    input("\nPress Enter to exit...")
      

################################################################################################################################

#funtion to hide or show extension of files within WINOS
REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
VALUE_NAME = "HideFileExt"

def get_current_setting():
    """Obtiene el estado actual de 'Mostrar extensiones de archivo'."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH, 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, VALUE_NAME)
        return value == 0  # Devuelve True si las extensiones están visibles, False si están ocultas
    except Exception as e:
        print("[Sorento] ❌ Error al leer la configuración:", e)
        return None

def refresh_explorer():
    """Usa un mensaje de Windows para actualizar el Explorador de archivos."""
    # Enviar mensaje a la ventana del Explorador de archivos para forzar actualización
    ctypes.windll.user32.PostMessageW(0xFFFF, 0x111, 41504, 0)
    print("[Sorento] 🔄 Explorador de archivos actualizado.")

def set_show_file_extensions(enable=True):
    """Modifica el registro de Windows para mostrar u ocultar las extensiones de archivo."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, VALUE_NAME, 0, winreg.REG_DWORD, 0 if enable else 1)

        print("\n[Sorento] ✅ Configuración aplicada correctamente.")
        
        # Actualizar el Explorador de archivos para aplicar cambios
        refresh_explorer()

        # Verificar si el cambio realmente se aplicó
        new_state = get_current_setting()
        if new_state == enable:
            print("[Sorento] ✅ Las extensiones de archivo ahora están visibles." if enable else "[Sorento] ❌ Las extensiones de archivo ahora están ocultas.")
        else:
            print("[Sorento] ⚠️ Parece que el cambio no se aplicó correctamente. Inténtalo manualmente.")

    except Exception as e:
        print("[Sorento] ❌ Error al cambiar la configuración:", e)
        
def toggle_file_extensions():
    """Function to show or hide file extensions in Windows."""
    print("\n🔹 Welcome to **Sorento** 🔹\nThis program allows you to show or hide file extensions in Windows.\n")

    current_state = get_current_setting()
    if current_state is not None:
        print(f"[Sorento] 📂 Current state: {'Extensions visible' if current_state else 'Extensions hidden'}")
    print("If you type NO, file extensions will be hide")
    response = input("[Sorento] Do you want to show file extensions? (yes/no): ").strip().lower()
    
    if response in ["yes"]:
        set_show_file_extensions(True)
    elif response in ["no"]:
        set_show_file_extensions(False)
    else:
        print("[Sorento] ❌ Invalid response. Type 'yes' or 'no'.")



####################

# Función para limpiar la pantalla
def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

# Menú principal
def menu():
    print("\033[36m=====================================================================================================================")
    print("\t\tHermes, the updater of Orion's software v 1.1.0")
    print("\t\tSoftware for upgrading and installing packages")
    print("\nThis software automatically downloads mbctools and BLAST+ and adds them to the PATH.")    
    print("Additionally, it installs Biopython and upgrades outdated Python modules.")
    
    print("\033[36m«Beyond mbctools: a set of plugins useful for further analysis in metabarcoding studies. ")
    print("2024, MOO-MILLAN JOEL ISRAEL» \033[36;1;3mGitHub: https://github.com/JOEL-ISRAEL-MOO-MILLAN\033[0m\033[36m")
    print("The first step is to choose step 1, then step 2,3,4, and optionally 5 and 6.")
    print("=====================================================================================================================\n\x1b[0m")


    print("\033[0;1;33m1. Check Execution Policy (Mandatory)")
    print("\n2. Run Hermes to install VSEARCH")
    print("\n3. Run Mercurio to install mbctools")
    print("\n4. Run Demeter to install packages for Orion program")
    print("\n5. Upgrade Packages")
    print("\n6. Show extensions files (mandatory)")
    print("\n7. Download and Configure BLAST")
    print("\n8. Change Execution Policy to 'Restricted'.\n\033[91m(CAUTION! run this option ONLY if you want to go back to the initial configuration of your computer and YOU DON'T WANT\nTO FURTHER RUN MBCTOOLS. If you still want to use mbctools, don't run this option)")
    print("\n\033[0;1;33m9. Exit")

    return input(titleStyle + "\nCHOOSE AN OPTION: ").strip()

############################################################################################################

import subprocess
import sys
import traceback

# Limpiar pantalla
def clear_screen():
    """Limpia la pantalla en Windows y Linux/Mac."""
    import os
    os.system('cls' if os.name == 'nt' else 'clear')

# Función para actualizar pip
def upgrade_pip():
    """Actualiza pip a la última versión."""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
        print("\nPip actualizado correctamente.\n")
    except Exception as e:
        print(f"\n\x1b[31mError al actualizar pip: {e}\x1b[0m\n")

# Función para verificar e instalar paquetes
def verify_installation(package_name, import_name=None):
    """Verifica si un paquete está instalado y lo instala si no está disponible."""
    try:
        import_name = import_name or package_name  # Usar el nombre del paquete si no hay alias
        __import__(import_name)
    except ImportError:
        print(f"\n\x1b[33mInstalando {package_name}...\x1b[0m\n")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
            print(f"\n\x1b[32m{package_name} instalado correctamente.\x1b[0m\n")
        except Exception as e:
            print(f"\n\x1b[31mError al instalar {package_name}: {e}\x1b[0m\n")

# Modo depuración
def debug_mode(func):
    """Función decoradora para capturar y mostrar errores detallados."""
    def wrapper(*args, **kwargs):
        try:
            func(*args, **kwargs)
        except Exception as e:
            print("\n\x1b[31mOcurrió un error inesperado:\x1b[0m")
            print(traceback.format_exc())  # Muestra el rastreo del error completo
            input("\n\x1b[36mPresiona Enter para salir...\x1b[0m")
    return wrapper

# Ejecutar Hermes
@debug_mode
def execute_hermes():
    """Función principal de Hermes."""
    clear_screen()
    print("_______________________________________________________________________________________________")
    print("|                                                                                             |")
    print("|                     ***Running Mercurio (v1.1.0), by Calaf***                                   |")
    print("|_____________________________________________________________________________________________|")
    print("\n\n")

    upgrade_pip()

    packages_to_check = [
        ('statsmodels', None), ('numpy', None), ('tqdm', None),
        ('pandas', None), ('openpyxl', None), ('requests', None),
        ('biopython', None), ('bs4', None)
    ]

    for package_name, package_import_name in packages_to_check:
        verify_installation(package_name, package_import_name)
    
    input("\n\x1b[36mPresiona Enter para continuar...\x1b[0m")
    clear_screen()



#####################################################################################################
#line for installing mbctools automatically
import subprocess
import time
from tqdm import tqdm

def installing_mbctools():
    """Instala mbctools con barra de progreso."""
    clear_screen()
    print("\n\x1b[34mInstalling mbctools, please be patient...\x1b[0m\n")
    
    try:
        import mbctools
        print("\n\x1b[32m✨ mbctools was installed successfully .\x1b[0m")
    except ImportError:
        print("\n\x1b[33mInstalling mbctools...\x1b[0m")
        process = subprocess.Popen(['pip', 'install', 'mbctools'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        for _ in tqdm(range(100), desc="Progreso", ascii=True):
            time.sleep(0.05)
            if process.poll() is not None:
                break
        
        process.communicate()

        if process.returncode == 0:
            print("\n\x1b[32mInstalación completada.\x1b[0m")
        else:
            print("\n\x1b[31mHubo un error en la instalación.\x1b[0m")

    input("\n\x1b[36mPress Enter to continue...\x1b[0m")
    clear_screen()

###########################################################################

# Actualizar paquetes
@debug_mode
def update_packages():
    """Actualiza pip y paquetes esenciales."""
    clear_screen()
    print("\n\x1b[34mActualizando pip y paquetes...\x1b[0m\n")
    
    upgrade_pip()

    packages_to_check = [
        ('tqdm', None), ('mbctools', None), ('pandas', None),
        ('openpyxl', None), ('requests', None), ('biopython', None), ('bs4', None)
    ]

    for package_name, package_import_name in packages_to_check:
        verify_installation(package_name, package_import_name)

    input("\n\x1b[36mPresiona Enter para continuar...\x1b[0m")
    clear_screen()



# Menú principal
if __name__ == "__main__":
    while True:
        clear_screen()
        option = menu()
        
        if option == '1':
            clear_screen()
            if check_execution_policy():
                print(successStyle + "\n✔  Execution policy checked " + normalStyle)
            else:
                print(errorStyle + "Could not set the execution policy. Exiting the program..." + normalStyle)
                break
            input("\n\x1b[36mPress enter to continue\x1b[0m")
        elif option == '2':
            install_vsearch()
        elif option == '3':
            installing_mbctools()
        elif option == '4':
            execute_hermes ()
        elif option == '5':
            update_packages()
        elif option == '6':
            clear_screen()
            toggle_file_extensions()
        elif option == '7':
            clear_screen()
            download_blast()
            input("\n\x1b[36mPress enter to continue\x1b[0m")
        elif option == '8':
            clear_screen()                                                
            input("\n\x1b[36mPress enter to continue\x1b[0m")
            change_execution_policy_to_restricted()

        elif option == '9':
            print(exitStyle + "Goodbye!" + normalStyle)
            print("If you have goy any question or want to report any bug, please, write to: etienne.waleckx@ird.fr")
            print("Sofware was designed by Joel Israel Moo-Millan joel.moo.millan@hotmail.com")
            break

        else:
            print(errorStyle + "Invalid option. Please try again." + normalStyle)

            
input("Press enter to close the program")
