import os
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, Scrollbar
import subprocess
from docx import Document
import PyPDF2

# Dossier par défaut
selected_folder = r"L:\3 - METTRE EN OEUVRE LA FORMATION\2 - PILOTER\CVs"

# Pour afficher seulement le dernier dossier
def update_folder_label():
    folder_label.config(text=f"Dossier sélectionné : {os.path.basename(selected_folder)}")

# Fonction pour extraire texte d'un fichier
def extract_text(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    text = ""
    try:
        if ext == ".txt":
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                text = f.read()
        elif ext == ".docx":
            doc = Document(file_path)
            text = "\n".join([p.text for p in doc.paragraphs])
        elif ext == ".pdf":
            with open(file_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                for page in reader.pages:
                    text += page.extract_text() or ""
    except Exception as e:
        print(f"Erreur lecture {file_path}: {e}")
    return text.lower()

# Fonction pour sélectionner un dossier
def choose_folder():
    global selected_folder
    folder = filedialog.askdirectory(title="Sélectionner le dossier contenant les CVs")
    if folder:
        selected_folder = folder
        update_folder_label()

# Fonction de recherche
def search_files():
    global selected_folder
    if not selected_folder:
        messagebox.showwarning("Alerte", "Veuillez sélectionner un dossier d'abord.")
        return

    keyword = entry.get().lower()
    if not keyword:
        messagebox.showwarning("Alerte", "Veuillez entrer un mot-clé.")
        return
    
    listbox.delete(0, tk.END)
    for root, _, files in os.walk(selected_folder):
        for file in files:
            if file.lower().endswith((".pdf", ".docx", ".txt")):
                file_path = os.path.join(root, file)
                if keyword in extract_text(file_path):
                    listbox.insert(tk.END, file_path)

# Fonction pour ouvrir le fichier sélectionné
def open_file(event):
    selected = listbox.get(listbox.curselection())
    if os.name == 'nt':  # Windows
        os.startfile(selected)
    elif os.name == 'posix':  # macOS / Linux
        subprocess.call(('xdg-open', selected))

# Interface Tkinter
root = tk.Tk()
root.title("Recherche mots-clés CV")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

tk.Label(frame, text="Mot-clé :").pack(side=tk.LEFT)
entry = tk.Entry(frame, width=30)
entry.pack(side=tk.LEFT, padx=5)
entry.bind("<Return>", lambda event: search_files())

tk.Button(frame, text="Chercher", command=search_files).pack(side=tk.LEFT)

# Bouton pour changer le dossier si besoin
tk.Button(root, text="Choisir dossier", command=choose_folder).pack(pady=5)
folder_label = tk.Label(root, text=f"Dossier sélectionné : {selected_folder}")
folder_label.pack()
update_folder_label()

# Liste des résultats
listbox_frame = tk.Frame(root)
listbox_frame.pack(padx=10, pady=10)
scrollbar = Scrollbar(listbox_frame)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

listbox = Listbox(listbox_frame, width=80, height=20, yscrollcommand=scrollbar.set)
listbox.pack()
listbox.bind("<Double-Button-1>", open_file)
scrollbar.config(command=listbox.yview)

root.mainloop()
