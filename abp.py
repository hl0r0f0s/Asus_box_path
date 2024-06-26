from pathlib import Path
from win32com.client import Dispatch
import getpass
import os
import click

print('')
print('░  ░░░░  ░░  ░░░░░░░░░      ░░░       ░░░░      ░░░        ░░░      ░░░░      ░░')
print('▒  ▒▒▒▒  ▒▒  ▒▒▒▒▒▒▒▒  ▒▒▒▒  ▒▒  ▒▒▒▒  ▒▒  ▒▒▒▒  ▒▒  ▒▒▒▒▒▒▒▒  ▒▒▒▒  ▒▒  ▒▒▒▒▒▒▒')
print('▓        ▓▓  ▓▓▓▓▓▓▓▓  ▓▓▓▓  ▓▓       ▓▓▓  ▓▓▓▓  ▓▓      ▓▓▓▓  ▓▓▓▓  ▓▓▓      ▓▓')
print('█  ████  ██  ████████  ████  ██  ███  ███  ████  ██  ████████  ████  ████████  █')
print('█  ████  ██        ███      ███  ████  ███      ███  █████████      ████      ██')                                                                               
print('')
print('░░      ░░░      ░░  ░░░░  ░░      ░░░░░░░       ░░░      ░░  ░░░░  ░░░░░░       ░░░      ░░        ░  ░░░░  ░')
print('▒  ▒▒▒▒  ▒  ▒▒▒▒▒▒▒  ▒▒▒▒  ▒  ▒▒▒▒▒▒▒▒▒▒▒▒  ▒▒▒▒  ▒  ▒▒▒▒  ▒▒  ▒▒  ▒▒▒▒▒▒▒  ▒▒▒▒  ▒  ▒▒▒▒  ▒▒▒▒  ▒▒▒▒  ▒▒▒▒  ▒')
print('▓  ▓▓▓▓  ▓▓      ▓▓  ▓▓▓▓  ▓▓      ▓▓▓▓▓▓▓       ▓▓  ▓▓▓▓  ▓▓▓    ▓▓▓▓▓▓▓▓       ▓▓  ▓▓▓▓  ▓▓▓▓  ▓▓▓▓        ▓')
print('█        ███████  █  ████  ███████  ██████  ████  █  ████  ██  ██  ███████  ███████        ████  ████  ████  █')
print('█  ████  ██      ███      ███      ███████       ███      ██  ████  ██████  ███████  ████  ████  ████  ████  █')                                                                                                                    
print('')

def create_shortcut(file_name: str, target: str, work_dir: str, arguments: str = ''):
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(file_name)
    shortcut.TargetPath = target
    shortcut.Arguments = arguments
    shortcut.WorkingDirectory = work_dir
    shortcut.save()

abs_file_name = r'C:\\Program Files\\ASUSTeKcomputer.Inc\\nhAsusStrix\\UserInterface\\nhAsusStrixSvc32.exe'
path = Path(abs_file_name)

user = getpass.getuser()

link = f"C:\\Users\\{user}\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup\\Asus_box_path.lnk"

create_shortcut(
    file_name= link,
    target=str(path),
    work_dir=str(path.parent),
    arguments='/start StrixControl',
)

print(f'[info] Shortcut create! Path: \n {link}')

os.startfile (link)

print(f'[info] Shortcut started! Finish!')

click.pause('[input] Press Enter to exit...')

