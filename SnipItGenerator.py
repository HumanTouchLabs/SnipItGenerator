import os
import re
import sys
import time
import logging
import pyperclip
import pygetwindow as gw
from pynput import keyboard
import ctypes
from ctypes import wintypes
import win32process
import win32api
import win32con
import pythoncom
import win32gui
import psutil
import win32com.client

# Configure logging
log_filename = 'SnipItGenerator.log'
logging.basicConfig(filename=log_filename, level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Add initial log entry
logging.info("Starting SnipItGenerator...")

def clean_filename_extension(filename, extension):
    invalid_chars = ['*', '/', '\\', '/*', '*/', '//']
    for char in invalid_chars:
        filename = filename.replace(char, '')
        extension = extension.replace(char, '')
    return filename.strip(), extension.strip()

def parse_filename_extension(line):
    logging.debug(f"Original line: {line}")
    reversed_line = line[::-1]
    logging.debug(f"Reversed line: {reversed_line}")
    
    if '*/' in reversed_line:
        parts = re.split(r'\s+|/|\\', reversed_line, 3)
        logging.debug(f"Parsed parts with '*/': {parts}")
        if len(parts) > 3:
            filename_part = parts[3][::-1]
            extension_part = parts[2][::-1]
        else:
            logging.debug("Filename or extension not properly extracted for '*/' case.")
            return None, None
    else:
        parts = re.split(r'\s+|/|\\', reversed_line, 2)
        logging.debug(f"Parsed parts: {parts}")
        if len(parts) > 1:
            filename_part = parts[1][::-1]
            extension_part = parts[0][::-1]
        else:
            logging.debug("Filename or extension not properly extracted.")
            return None, None
    
    filename_parts = filename_part.rsplit('.', 1)
    if len(filename_parts) == 2:
        filename, extension = filename_parts
    else:
        filename = filename_part
        extension = extension_part

    filename, extension = clean_filename_extension(filename, extension)
    logging.debug(f"Cleaned filename: {filename}, Cleaned extension: {extension}")

    if filename and extension:
        return filename, extension
    else:
        logging.debug("Invalid filename or extension after cleaning.")
        return None, None

def get_directory_from_mouse_cursor():
    pythoncom.CoInitialize()
    try:
        hwnd = win32gui.WindowFromPoint(win32api.GetCursorPos())
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)

        if process.name().lower() == 'explorer.exe':
            shell = win32com.client.Dispatch("Shell.Application")
            while hwnd:
                for window in shell.Windows():
                    if window.HWND == hwnd:
                        directory = window.Document.Folder.Self.Path
                        logging.debug(f"Directory from SHGetPathFromIDList: {directory}")
                        return directory
                hwnd = win32gui.GetParent(hwnd)
    except Exception as e:
        logging.error(f'Error getting directory from mouse cursor: {e}')
    finally:
        pythoncom.CoUninitialize()
    return None

def on_clipboard_change():
    try:
        clipboard_content = pyperclip.paste()
        logging.info("Clipboard content detected.")
        lines = clipboard_content.split('\n')
        if lines:
            logging.info(f"Comment line detected: {lines[0]}")
            filename, extension = parse_filename_extension(lines[0])
            if filename and extension:
                logging.debug(f"Extracted filename: {filename}, Extracted extension: {extension}")
                directory = get_directory_from_mouse_cursor()
                if not directory or directory.lower() == 'c:\\windows':
                    directory = os.path.join(os.path.expanduser("~"), "Desktop")
                file_path = os.path.join(directory, f"{filename}.{extension}")
                logging.info(f"Creating file with name: {file_path}")
                with open(file_path, 'w') as f:
                    f.write(clipboard_content)
                logging.info(f"File created: {file_path}")
            else:
                logging.error("Invalid filename or extension.")
    except Exception as e:
        logging.error(f"Error: {e}")

def on_activate_v():
    logging.info("Ctrl+V detected.")
    on_clipboard_change()

def for_canonical(f):
    return lambda k: f(l.canonical(k))

def main():
    logging.info("Starting SnipItGenerator...")
    with keyboard.GlobalHotKeys({
            '<ctrl>+v': on_activate_v}) as h:
        h.join()

if __name__ == "__main__":
    main()
