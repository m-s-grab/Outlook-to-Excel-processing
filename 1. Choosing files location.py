import os
import json
import tkinter as tk
from tkinter import filedialog
import sys

# Configuration file path
def get_config_path():
    if getattr(sys, 'frozen', False):  # Running as an executable
        return os.path.join(os.path.dirname(sys.executable), "config.json")
    else:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

CONFIG_FILE = get_config_path()

# Hide the main Tkinter window
root = tk.Tk()
root.withdraw()

# Load config
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

# Save config
def save_config(process_dir=None, supplier_file=None):
    config = load_config()
    if process_dir:
        config["processing_location"] = process_dir
    if supplier_file:
        config["supplier_file_location"] = supplier_file
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2)

# Select processing folder
def select_processing_location():
    config = load_config()
    current = config.get("processing_location")

    if current and os.path.isdir(current):
        print(f"\n📁 Current Data folder:\n➡️  {current}")
        if input("Do you want to change it? (y/n): ").strip().lower() != "y":
            return current

    while True:
        print("Select the Data folder in the pop-up dialog")
        new = filedialog.askdirectory(title="Select Processing Folder")
        if os.path.isdir(new):
            print(f"\n➡️ Selected: {new}")
            if input("Is this path correct? (y/n): ").strip().lower() == "y":
                save_config(process_dir=new)
            return new
        print("❌ Invalid folder.")

# Select supplier master file
def select_supplier_file():
    config = load_config()
    current = config.get("supplier_file_location")

    if current and os.path.isfile(current):
        print(f"\n📄 Current supplier master file:\n➡️  {current}")
        if input("Do you want to change it? (y/n): ").strip().lower() != "y":
            return current

    while True:
        print("Select the supplier master file in the pop-up dialog")
        new = filedialog.askopenfilename(
            title="Select Supplier Master File (.xlsx/.xlsm)",
            filetypes=[("Excel files", "*.xlsx *.xlsm")]
        )
        if os.path.isfile(new):
            print(f"\n➡️ Selected: {new}")
            if input("Is this path correct? (y/n): ").strip().lower() == "y":
                save_config(supplier_file=new)
            return new
        print("❌ Invalid file.")

# Reset configuration
def reset_config():
    if os.path.exists(CONFIG_FILE):
        os.remove(CONFIG_FILE)
    print("🛠️ Configuration has been reset!")

# Main script logic
def main():
    while True:
        processing_path = select_processing_location()
        supplier_path = select_supplier_file()

        print("\n📂 Configuration complete!")

        print("\n📋 What would you like to do next?")
        print("1. Select locations again")
        print("2. Reset configuration and exit")
        print("3. Exit")

        choice = input("Choose (1-3): ").strip()

        if choice == "1":
            continue
        elif choice == "2":
            reset_config()
            break
        elif choice == "3":
            print("Finished.")
            break
        else:
            print("❌ Invalid selection.")

if __name__ == "__main__":
    main()
