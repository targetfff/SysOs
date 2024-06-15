import winreg
import os
from pyuac import main_requires_admin


@main_requires_admin
def main():
    extensions = ['pdf', 'txt', 'docx', 'mp3', 'wav', 'ogg', 'avi', '3ga', '4xm', 'aac',
                  'htk', 'g722', 'weba', 'opus', 'mov', 'vmv', 'gif', 'm4a', 'flac', 'mp4',
                  'mkv', 'h261', 'h262', 'h263', 'h265', 'dct', 'cak', 'av1', 'xvv', 'webm',
                  'paf', 'p64', 'pmp', 'jpeg', 'png', 'img', 'bmp', 'gif', 'pfm', 'spider',
                  'tiff', 'blp', 'icns', 'ico', 'jpg', 'tga', 'webp', 'dds', 'dib', 'eps',
                  'im', 'msp', 'pcx', 'xbm', 'jpg']
    exe_main = 'SysOs.exe'
    base_file_name = os.path.basename(exe_main)
    absolute_path = os.path.abspath(base_file_name)
    path = winreg.HKEY_CLASSES_ROOT
    system_file_associations = winreg.OpenKeyEx(path, r"SystemFileAssociations\\",
                                                0, winreg.KEY_ALL_ACCESS)
    for extension in extensions:
        sub_key = rf'.{extension}\\'
        try:
            ext_key = winreg.OpenKeyEx(system_file_associations, sub_key, 0, winreg.KEY_ALL_ACCESS)
        except:
            ext_key = winreg.CreateKey(system_file_associations, sub_key)
        try:
            shell = winreg.OpenKeyEx(ext_key, r'Shell\\', 0, winreg.KEY_ALL_ACCESS)
        except:
            shell = winreg.CreateKey(ext_key, "Shell")
        SysOs = winreg.CreateKey(shell, 'Convert with SysOs')
        winreg.SetValueEx(SysOs, 'Icon', 0, winreg.REG_SZ,
                          rf'"{absolute_path}"')
        command = winreg.CreateKey(SysOs, 'command')
        winreg.SetValueEx(command, None, 0, winreg.REG_SZ,
                          rf'"{absolute_path}" "%1"')
        if command:
            winreg.CloseKey(command)
        if SysOs:
            winreg.CloseKey(SysOs)
        if shell:
            winreg.CloseKey(shell)
        if ext_key:
            winreg.CloseKey(ext_key)


if __name__ == '__main__':
    main()
