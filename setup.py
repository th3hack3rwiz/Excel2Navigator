import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# List of packages to install
packages = [
    "requests",
    "pandas",
    "openpyxl",
    "matplotlib",
    "pyfiglet"
]

for package in packages:
    try:
        install(package)
        print(f"{package} installed successfully!")
    except subprocess.CalledProcessError:
        print(f"Failed to install {package}")

print("Requirements for Excel2Navigator installed successfully!")
