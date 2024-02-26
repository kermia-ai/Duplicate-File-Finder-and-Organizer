import os
import hashlib
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from openpyxl import Workbook
import datetime

def hash_file(filepath):
    """Calcule le hash MD5 d'un fichier."""
    hasher = hashlib.md5()
    with open(filepath, 'rb') as file:
        buf = file.read(65536)  # Lire le fichier par blocs de 64ko
        while len(buf) > 0:
            hasher.update(buf)
            buf = file.read(65536)
    return hasher.hexdigest()

def write_duplicates_to_excel(duplicates, file_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["Chemin du fichier", "Taille (Mo)", "Nombre de doublons", "Type de fichier", "Dernière modification"])

    for paths in duplicates.values():
        file_count = len(paths)
        file_size = os.path.getsize(paths[0]) / (1024*1024)  # Converti en Mo
        for path in paths:
            file_extension = os.path.splitext(path)[1]
            modification_time = datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime('%Y-%m-%d %H:%M:%S')
            ws.append([path, f"{file_size:.5f}", file_count, file_extension, modification_time])

    wb.save(file_path)

def find_duplicates(start_dir):
    """Trouve et enregistre les fichiers en double dans le répertoire donné et ses sous-répertoires."""
    hashes = defaultdict(list)
    for dir_name, _, file_list in os.walk(start_dir):
        print(f"Scanning {dir_name}...")
        for filename in file_list:
            filepath = os.path.join(dir_name, filename)
            try:
                file_hash = hash_file(filepath)
                hashes[file_hash].append(filepath)
            except (OSError,):
                pass  # Ignorer les fichiers qui ne peuvent pas être ouverts

    # Identifier les fichiers en double
    duplicates = {hash: paths for hash, paths in hashes.items() if len(paths) > 1}

    if duplicates:
        print("Fichiers en double trouvés :")
        # output_file = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

        if output_file:
            # write_duplicates_to_file(duplicates, output_file)
            write_duplicates_to_excel(duplicates, output_file)
            print(f"Les doublons ont été enregistrés dans {output_file}")
            messagebox.showinfo("Terminé", f"Les doublons ont été enregistrés dans {output_file}")
        else:
            print("Opération annulée.")
    else:
        print("Aucun fichier en double trouvé.")
        messagebox.showinfo("Résultat", "Aucun fichier en double trouvé.")

def browse_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        find_duplicates(folder_selected)

app = tk.Tk()
app.title('Trouveur de Fichiers en Double')
app.geometry('400x150')

label = tk.Label(app, text="Sélectionnez un dossier pour rechercher des fichiers en double :", wraplength=400)
label.pack(pady=10)

browse_button = tk.Button(app, text="Parcourir", command=browse_folder)
browse_button.pack(pady=5)

close_button = tk.Button(app, text="Fermer", command=app.destroy)
close_button.pack(pady=5)

app.mainloop()
