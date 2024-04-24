import subprocess
import sys

def install_and_import(package):
    try:
        __import__(package)
        # print(f"{package} is already installed.")
    except ImportError:
        # print(f"{package} not found, installing using pip...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        try:
            __import__(package)
            # print(f"{package} has been successfully installed.")
        except ImportError:
            print(f"Failed to install {package}. Please install it manually.")

# List of required packages
required_packages = ['pandas', 'pyautogui', 'pyperclip', 'keyboard', 'openpyxl']

# Attempt to import or install each package
for package in required_packages:
    install_and_import(package)
print("All dependencies installed")
