import os
import pyperclip
import keyboard
import pystray
from PIL import Image, ImageDraw
import threading
import logging
import re
import win32gui
import win32api
import win32process
import psutil
import pythoncom
import win32com.client

# Configure logging to create a new log file every time the script is started
if os.path.exists('SnipItGenerator.log'):
    os.remove('SnipItGenerator.log')
logging.basicConfig(filename='SnipItGenerator.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Dictionary of supported file extensions and their identifying patterns
SUPPORTED_EXTENSIONS = {
    'js': r'\bfunction\b|\bconst\b|\blet\b|\bvar\b',
    'go': r'\bpackage main\b',
    'py': r'\bdef\b',
    'java': r'\bpublic class\b',
    'c': r'#include\b',
    'cpp': r'#include\b',
    'cs': r'\busing System;\b',
    'rb': r'\bdef\b',
    'php': r'<\?php\b',
    'html': r'<html\b',
    'css': r'\bbody\s*{',
    'sql': r'\bSELECT\b',
    'sh': r'#!/bin/bash\b',
    'xml': r'<\?xml\b',
    'swift': r'\bimport Swift\b',
    'kt': r'\bfun main\b',
    'r': r'# R script\b',
    'ts': r'\bimport {|\bexport\b',
    'h': r'#include\b',
    'json': r'^\s*[{[]',
    'yaml': r'^\s*---',
    'yml': r'^\s*---',
    'md': r'^#',
    'ini': r'^\[',
    'bat': r'@echo\b',
    'ps1': r'^\s*<#\s*PSScriptInfo',
    'vbs': r'^\s*<\?\s*VBScript',
    'pl': r'#!/usr/bin/perl\b',
    'm': r'\bfunction\b',
    'coffee': r'# CoffeeScript\b',
    'gitignore': r'^#',
    'gitattributes': r'^merge=',
    'p4': r'\bPerforce\b',
    'p4ignore': r'^#',
    'uasset': r'UE4Asset\b',
    'uproject': r'^\s*{',
    'meta': r'\bfileFormatVersion\b',
    'unity': r'\bm_Script\b',
    'tres': r'^\[gd_resource\b',
    'tscn': r'^\[gd_scene\b',
}

# Get the directory from the mouse cursor position
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
                        return directory
                hwnd = win32gui.GetParent(hwnd)
    except Exception as e:
        logging.error(f'Error getting directory from mouse cursor: {e}')
    finally:
        pythoncom.CoUninitialize()
    return None

# Sanitize the filename to remove invalid characters
def sanitize_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "", filename)

# Reverse the first line and extract the filename and extension
def reverse_and_extract_filename(line):
    reversed_line = line[::-1]
    parts = re.split(r'[ \\\/]', reversed_line, 1)
    if len(parts) == 2:
        ext_and_filename = parts[0].split('.', 1)
        if len(ext_and_filename) == 2:
            extension = ext_and_filename[0][::-1].strip()
            filename = ext_and_filename[1][::-1].strip()
            return filename, extension
    return None, None

# Create a file with the given name and content in the specified directory
def create_file(filename, content, directory):
    try:
        filepath = os.path.join(directory, filename)
        with open(filepath, 'w') as file:
            file.write(content)
        logging.info(f'File created: {filepath}')
    except Exception as e:
        logging.error(f'Failed to create file: {e}')

# Check the clipboard content and create a file based on the extracted information
def check_clipboard_and_create_file():
    clipboard_content = pyperclip.paste()
    if clipboard_content:
        first_line = clipboard_content.split('\n')[0]
        base_filename, extension = reverse_and_extract_filename(first_line)
        
        if base_filename and extension:
            extension_found = extension in SUPPORTED_EXTENSIONS
            logging.info(f'Extension found in supported list: {extension_found}')
            logging.info(f'Comment line detected: {base_filename}.{extension}')
            directory = get_directory_from_mouse_cursor()
            if directory:
                filename = f'{base_filename}.{extension}'
                logging.info(f'Creating file with name: {filename} in directory: {directory}')
                create_file(filename, clipboard_content, directory)
            else:
                logging.error('Directory detection failed.')
            return

        for ext, pattern in SUPPORTED_EXTENSIONS.items():
            if re.search(pattern, clipboard_content, re.IGNORECASE):
                directory = get_directory_from_mouse_cursor()
                if directory:
                    filename = f'file.{ext}'
                    logging.info(f'Creating file with name: {filename} in directory: {directory}')
                    create_file(filename, clipboard_content, directory)
                else:
                    logging.error('Directory detection failed.')
                return

        if base_filename and extension:
            directory = get_directory_from_mouse_cursor()
            if directory:
                filename = f'{base_filename}.{extension}'
                logging.info(f'Creating file with name: {filename} in directory: {directory}')
                create_file(filename, clipboard_content, directory)
            else:
                logging.error('Directory detection failed.')

# Handle Ctrl+V event to trigger the file creation
def on_key_event(event):
    if event.name == 'v' and keyboard.is_pressed('ctrl'):
        check_clipboard_and_create_file()

# Create the system tray icon
def create_image():
    icon_path = os.path.join(os.path.dirname(__file__), 'htllogo.ico')
    if os.path.exists(icon_path):
        return Image.open(icon_path)
    width, height = 64, 64
    image = Image.new('RGB', (width, height), 'white')
    dc = ImageDraw.Draw(image)
    dc.rectangle((width // 2, 0, width, height // 2), fill='black')
    dc.rectangle((0, height // 2, width // 2, height), fill='black')
    return image

# Handle quitting the application from the system tray
def on_quit(icon, item):
    logging.info('Quitting the application.')
    icon.stop()
    os._exit(0)

# Set up the system tray icon
def setup_tray_icon():
    icon = pystray.Icon("SnipItGenerator")
    icon.icon = create_image()
    icon.title = "SnipItGenerator"
    icon.menu = pystray.Menu(pystray.MenuItem('Quit', on_quit))
    icon.run()

if __name__ == "__main__":
    logging.info('Starting SnipItGenerator...')
    print("Monitoring for Ctrl+V... Click the tray icon to quit.")

    tray_thread = threading.Thread(target=setup_tray_icon)
    tray_thread.daemon = True
    tray_thread.start()

    keyboard.hook(on_key_event)

    keyboard.wait()
