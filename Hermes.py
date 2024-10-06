import os
import subprocess
import zipfile
import urllib.request
import sys
import importlib
import tarfile
import warnings
import requests  # Importa después de verificar
from tqdm import tqdm  # Importa después de verificar

# añadir una función que habilita visualizar las extensiones de los archivos en WIN
# Definición de estilos para impresión en consola
errorStyle = "\033[91m"
warningStyle = "\033[93m"
normalStyle = "\033[0m"
titleStyle = "\033[34;1;3m"
promptStyle = "\033[96m"
successStyle = "\033[92m"
runStyle = "\x1B[38;2;255;128;0m"
citestyle = "\x1B[38;2;17;245;120m"
exitStyle = "\033[34;1;3m"

# Función para instalar paquetes si no están presentesz
def install_and_import(package_name):
    print("Updating and installing some libraries needed for this program, please be patient...")
    try:
        importlib.import_module(package_name)
        print(successStyle + f"{package_name} is already installed ✔" + normalStyle)
    except ImportError:
        #print(f"{package_name} is not installed. Installing now...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', package_name])
        #print(successStyle + f"{package_name} installed successfully ✔" + normalStyle)

# Verificar e instalar paquetes necesarios
def verify_and_install_packages():
    install_and_import('requests')
    install_and_import('tqdm')

# Llamada inicial para asegurarse de que los paquetes estén disponibles
verify_and_install_packages()



def download_and_extract(url, destination):
    try:
        # Prepare the file name and download the file
        print(titleStyle + "\n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        print("For more information see: \nhttps://blast.ncbi.nlm.nih.gov/doc/blast-help/downloadblastdata.html")
        print("\n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        print(warningStyle + "Establishing connection with NCBI page")
        print(successStyle + "\nDownloading BLAST+, be patient and make sure you have got a stable internet connection...\n" + promptStyle )
        
        
        file_name = os.path.join(destination, url.split('/')[-1])

        with requests.get(url, stream=True) as response:
            response.raise_for_status()  # Lanzar un error para respuestas incorrectas
            total_size = int(response.headers.get('Content-Length', 0))
            block_size = 1024  # Tamaño del bloque para la barra de progreso
            tqdm_bar = tqdm(total=total_size, 
                            unit='B', 
                            unit_scale=True, 
                            desc= runStyle + 'Downloading BLAST+' )

            with open(file_name, 'wb') as f:
                for chunk in response.iter_content(chunk_size=block_size):
                    f.write(chunk)
                    tqdm_bar.update(len(chunk))
            tqdm_bar.close()

        print(successStyle + f"\nBLAST+ was stored in: {destination} ✔")


        # Ignorar las advertencias de deprecación
        warnings.simplefilter(action='ignore', category=DeprecationWarning)
        
        # Extraer el archivo .tar.gz
        print(runStyle + "Unzipping files...")
        if tarfile.is_tarfile(file_name):
            with tarfile.open(file_name, "r:gz") as tar:
                def is_within_directory(directory, target):
                    abs_directory = os.path.abspath(directory)
                    abs_target = os.path.abspath(os.path.join(directory, target))
                    return os.path.commonpath([abs_directory]) == os.path.commonpath([abs_directory, abs_target])

                def safe_extract(tar, path=".", members=None):
                    for member in tar.getmembers():
                        if not is_within_directory(path, member.name):
                            raise Exception("La ruta del archivo está fuera del directorio de destino.")
                    tar.extractall(path, members)

                safe_extract(tar, destination)

            print(successStyle + f"Archivo descomprimido en: {destination} ✔")
        else:
            print(warningStyle + "El archivo descargado no es un .tar.gz válido.")
            return None

    except Exception as e:
        print(warningStyle + f"Error durante la descarga o descompresión: {e}")
        return None

    return file_name

def find_bin_path(root_directory):
    # Buscar la carpeta 'bin' recursivamente en la estructura de directorios
    for root, dirs, files in os.walk(root_directory):
        if 'bin' in dirs:
            bin_path = os.path.join(root, 'bin')
            print(successStyle + f"\nFolder 'bin' was successfully found in: {bin_path} ✔")
            return bin_path
    
    print(warningStyle + "No se encontró la carpeta 'bin' en la ruta esperada.")
    return None

def add_blast_to_path_permanently(blast_path):
    try:
        current_path = os.environ['PATH']
        
        if blast_path not in current_path:
            # Obtener el PATH del usuario actual utilizando PowerShell
            result = subprocess.run(['powershell', '[System.Environment]::GetEnvironmentVariable("Path", "User")'], capture_output=True, text=True)
            user_path = result.stdout.strip()

            # Verificar si la ruta ya está en el PATH del usuario
            if blast_path not in user_path:
                new_path = user_path + ";" + blast_path
                # Añadir la ruta al PATH del usuario utilizando PowerShell
                subprocess.run(['powershell', f'[System.Environment]::SetEnvironmentVariable("Path", "{new_path}", "User")'])
                print(successStyle + f"\n{blast_path} was added successfully to the system PATH  ✔")
                print(warningStyle + "►\nCheck if the path was successfully added to the environmental variables◄♫♫♪")
            else:
                print(warningStyle + "La ruta de blast ya está en el PATH del sistema.")
        else:
            print(warningStyle + "La ruta de blast ya está en el PATH de la sesión actual.")
    except Exception as e:
        print(errorStyle + f"Error al añadir la ruta de blast al PATH del sistema: {e}")

def download_blast():
    url = "https://ftp.ncbi.nlm.nih.gov/blast/executables/blast+/LATEST/ncbi-blast-2.16.0+-x64-win64.tar.gz"
    destination = os.getcwd()  # Directorio actual donde se ejecuta el script
    tar_file = download_and_extract(url, destination)

    if tar_file:
        # Buscar la carpeta 'bin' dentro del directorio descomprimido
        bin_path = find_bin_path(destination)
        if bin_path:
            add_blast_to_path_permanently(bin_path)





# Función para verificar y modificar la política de ejecución de PowerShell
def check_execution_policy():
    try:
        print("=====================================")
        print("       Starting the process")
        print("=====================================")
        load_module = subprocess.run(['powershell', 'Import-Module', 'Microsoft.PowerShell.Security'], capture_output=True, text=True)
        if load_module.returncode != 0:
            print(f"Error loading module Microsoft.PowerShell.Security: {load_module.stderr}")
            return False

        result = subprocess.run(['powershell', 'Get-ExecutionPolicy'], capture_output=True, text=True)
        policy = result.stdout.strip()

        if policy != 'Unrestricted':
            print(f"Current execution policy: {policy}")
            user_response = input(runStyle + "\nWould you like 'Hermes' to change the policy to 'Unrestricted'? (yes/no): ").strip().lower()

            if user_response == 'yes':
                print("Attempting to change the execution policy to 'Unrestricted'...")
                change_policy = subprocess.run(['powershell', 'Set-ExecutionPolicy', 'Unrestricted', '-Scope', 'CurrentUser', '-Force'], capture_output=True, text=True)

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
            print("\n✔  Execution policy is already set to 'Unrestricted'! ")
            return True
    except Exception as e:
        print(f"Error checking or changing execution policy: {e}")
        return False

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

# Función para verificar la instalación de paquetes
def verify_installation(package_name, package_import_name=None):
    if package_import_name is None:
        package_import_name = package_name
    try:
        importlib.import_module(package_import_name)
        print(successStyle + f"{package_name} is already installed ✔" + normalStyle)
    except ImportError:
        print(f"{package_name} is not installed. Installing now...")
        install_package(package_name)

# Función para cambiar la política de ejecución a 'Restricted'
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


# Función para descargar y extraer VSEARCH
def download_vsearch(url, extract_to = "vsearch"):
    try:
        os.makedirs(extract_to, exist_ok=True)
        zip_file = os.path.join(extract_to, "vsearch.zip")
        print(f"Downloading vsearch from {url}...")
        urllib.request.urlretrieve(url, zip_file)
        print(successStyle + "Download completed ✔" + normalStyle)
        
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            print(f"Extracting to {extract_to}...")
            zip_ref.extractall(extract_to)
        print(successStyle + "Extraction completed ✔" + normalStyle)
        
        vsearch_dir = os.path.join(extract_to, "vsearch-2.28.1-win-x86_64", "bin")
        vsearch_exe = os.path.join(vsearch_dir, "vsearch.exe")
        
        if (os.path.exists(vsearch_exe)):
            print(successStyle + f"vsearch.exe found at {vsearch_exe}" + normalStyle)
            return vsearch_dir
        else:
            print(warningStyle + f"vsearch.exe not found in {vsearch_dir}" + normalStyle)
            return None
    except Exception as e:
        print(warningStyle + f"Error downloading or extracting vsearch: {e}" + normalStyle)
        return None

# Función para añadir VSEARCH al PATH permanentemente
def add_vsearch_to_path_permanently(vsearch_path):
    try:
        current_path = os.environ['PATH']
        
        if vsearch_path not in current_path:
            result = subprocess.run(['powershell', '[System.Environment]::GetEnvironmentVariable("Path", "User")'], capture_output=True, text=True)
            user_path = result.stdout.strip()

            if vsearch_path not in user_path:
                new_path = user_path + ";" + vsearch_path
                subprocess.run(['powershell', f'[System.Environment]::SetEnvironmentVariable("Path", "{new_path}", "User")'])
                print(successStyle + f"{vsearch_path} permanently added to the system PATH ✔" + normalStyle)
            else:
                print(successStyle + "vsearch path is already in the system PATH." + normalStyle)
        else:
            print(successStyle + "The vsearch path is already in the session PATH." + normalStyle)
    except Exception as e:
        print(warningStyle + f"Error adding vsearch path to system PATH: {e}" + normalStyle)


# Función para limpiar la pantalla
def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

# Menú principal
def menu():
    print("\033[36m=====================================================================================================================")
    print("\t\tHermes, the updater of Orion's software")
    print("\t\tSoftware for upgrading and installing packages")
    print("\nThis software automatically downloads mbctools and BLAST+ and adds them to the PATH.")    
    print("Additionally, it installs Biopython and upgrades outdated Python modules.")
    
    print("\033[36m«Beyond mbctools: a set of plugins useful for further analysis in metabarcoding studies. ")
    print("2024, MOO-MILLAN JOEL ISRAEL» \033[36;1;3mGitHub: https://github.com/JOEL-ISRAEL-MOO-MILLAN\033[0m\033[36m")
    print("The first step is to choose number 1, then number 2, and optionally 3 and 4.")
    print("=====================================================================================================================\n\x1b[0m")


    print("\033[0;1;33m1. Check Execution Policy (Mandatory)")
    print("2. Run Hermes")
    print("3. Upgrade Packages")
    print("4. Download and Configure BLAST")
    print("5. Change Execution Policy to 'Restricted'.\n\033[91m(CAUTION! run this option ONLY if you want to go back to the initial configuration of your computer and YOU DON'T WANT\nTO FURTHER RUN MBCTOOLS. If you still want to use mbctools, don't run this option)")
    print("\033[0;1;33m6. Exit")

    return input("\nChoose an option: ").strip()

# Ejecutar Hermes
def execute_hermes():
    clear_screen()
    print(titleStyle + "________________________________________________________________________________________________")
    print("|                                                                                               |")
    print("|                     ***Running Hermes (v1.0), by Calaf***                                     |")
    print("|_______________________________________________________________________________________________|")
    print("\n\n")
    upgrade_pip()

    packages_to_check = [ ('stats', None),  ('statsmodels', None), 
        ('tdqm', None), 
        ('numpy', None), 
        ('tqdm', None),
        ('mbctools', None),
        ('pandas', None), 
        ('product', None), 
        ('openpyxl', None), 
        ('math', None), 
        ('xml.etree.ElementTree', 'xml.etree.ElementTree'), 
        ('openpyxl.styles', 'openpyxl.styles'), 
        ('requests', None), 
        ('biopython', None)
    ]

    for package_name, package_import_name in packages_to_check:
        verify_installation(package_name, package_import_name)

    vsearch_url = "https://github.com/torognes/vsearch/releases/download/v2.28.1/vsearch-2.28.1-win-x86_64.zip"
    
    current_directory = os.getcwd()
    
    vsearch_path = download_vsearch(vsearch_url, current_directory)

    if vsearch_path:
        add_vsearch_to_path_permanently(vsearch_path)
    else:
        print(errorStyle + "vsearch.exe not found, cannot add to PATH." + normalStyle) 

    input("\n\x1b[36mPress enter to continue\x1b[0m")
    clear_screen()

# Actualizar paquetes
def update_packages():
    clear_screen()
    print(promptStyle + "\nUpgrading pip and packages...\n" + normalStyle)
    
    upgrade_pip()

    packages_to_check = [
        ('tqdm', None),
        ('mbctools', None),
        ('pandas', None), 
        ('product', None), 
        ('openpyxl', None), 
        ('math', None), 
        ('xml.etree.ElementTree', 'xml.etree.ElementTree'), 
        ('openpyxl.styles', 'openpyxl.styles'), 
        ('requests', None), 
        ('biopython', None)
    ]

    for package_name, package_import_name in packages_to_check:
        verify_installation(package_name, package_import_name)

    input("\n\x1b[36mPress enter to continue\x1b[0m")
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
            execute_hermes()

        elif option == '3':
            update_packages()

        elif option == '4':
            clear_screen()
            download_blast()
            input("\n\x1b[36mPress enter to continue\x1b[0m")

        elif option == '5':
            clear_screen()
                                                                        
            input("\n\x1b[36mPress enter to continue\x1b[0m")
            change_execution_policy_to_restricted()
        elif option == '6':
            print(exitStyle + "Goodbye!" + normalStyle)
            print("If you have goy any question or want to report any bug, please, write to: etienne.waleckx@ird.fr")
            print("Sofware was designed by Joel Israel Moo-Millan joel.moo.millan@hotmail.com")
            break

        else:
            print(errorStyle + "Invalid option. Please try again." + normalStyle)

            
input("Press enter to close the program")