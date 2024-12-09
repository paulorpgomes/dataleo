import os
import re
import pandas as pd
import chardet
from pymongo import MongoClient
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Funções utilitárias
def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read())
        return result['encoding']

def extract_date_from_filename(filename):
    month_map = {
        'jan': '01', 'fev': '02', 'mar': '03', 'abr': '04',
        'mai': '05', 'jun': '06', 'jul': '07', 'ago': '08',
        'set': '09', 'out': '10', 'nov': '11', 'dez': '12'
    }
    date_pattern = r'([a-zA-Z]{3})(\d{2})'
    match = re.search(date_pattern, filename.lower())
    if match:
        month_str, year_str = match.groups()
        month_num = month_map.get(month_str.lower())
        if month_num:
            try:
                date_obj = datetime.strptime(f"{year_str}-{month_num}", "%y-%m")
                return date_obj.strftime("%Y-%m")
            except ValueError:
                return None
    return None

def process_file(file_path, gestora):
    file_extension = os.path.splitext(file_path)[1].lower()
    encoding = detect_encoding(file_path)

    if file_extension == '.csv':
        data = pd.read_csv(file_path, encoding=encoding)
    elif file_extension == '.json':
        data = pd.read_json(file_path, encoding=encoding)
    elif file_extension == '.xls':
        data = pd.read_excel(file_path, engine='xlrd')
    elif file_extension == '.xlsx':
        data = pd.read_excel(file_path, engine='openpyxl')
    elif file_extension == '.txt':
        data = pd.read_csv(file_path, delimiter='\t', encoding=encoding)
    else:
        raise ValueError("Formato de arquivo não suportado")
    
    filter_settings = {
        "Acadian": {"rows": (6, 12), "columns": [0, 1]},
        "Colchester": {"rows": (15, 26), "columns": [10, 11]},
        "Fundamenta": {"rows": (4, 10), "columns": [0, 2]},
        "Lord Abett": {"rows": (6, 12), "columns": [0, 1]},
        "Man": {"rows": (2, 20), "columns": [0, 1]},
        "Oaktree": {"rows": (3, 8), "columns": [0, 1]},
        "Pearl Diver": {"rows": (7, 14), "columns": [0, 1]},
        "Zeno": {"rows": (1, 10), "columns": [1, 2]}
    }
    
    if gestora not in filter_settings:
        raise ValueError(f"Gestora {gestora} não possui configuração de filtragem definida.")
    
    settings = filter_settings[gestora]
    start_row, end_row = settings["rows"]
    columns = settings["columns"]
    
    filtered_data = data.iloc[start_row:end_row, columns]
    filtered_data.columns = ['Metric', 'Value']
    return filtered_data.dropna()

# MongoDB Configuração
client = MongoClient('mongodb://localhost:27017/')
db = client['projetoLeo']

def insert_filtered_data(collection_name, dataframe, filename):
    records = dataframe.to_dict(orient='records')
    document_date = extract_date_from_filename(filename)
    collection = db[collection_name]
    document = {
        "_id": filename,
        "data_documento": document_date,
        "dados": records
    }
    collection.replace_one({"_id": filename}, document, upsert=True)

def process_directory(directory_path, gestora):
    for filename in os.listdir(directory_path):
        file_path = os.path.join(directory_path, filename)
        if os.path.isfile(file_path):
            try:
                data = process_file(file_path, gestora)
                if not data.empty:
                    collection_name = gestora.lower().replace(" ", "_") + "_filtrados"
                    insert_filtered_data(collection_name, data, filename)
            except Exception as e:
                print(f"Erro ao processar {filename}: {e}")

# Interface gráfica
def create_ui():
    def selecionar_diretorio():
        directory = filedialog.askdirectory()
        if directory:
            entry_diretorio.delete(0, tk.END)
            entry_diretorio.insert(0, directory)

    def processar():
        directory = entry_diretorio.get()
        gestora = combo_gestoras.get()
        if not directory or not os.path.isdir(directory):
            messagebox.showwarning("Atenção", "Por favor, selecione um diretório válido.")
            return
        if not gestora:
            messagebox.showwarning("Atenção", "Por favor, selecione uma gestora.")
            return
        try:
            process_directory(directory, gestora)
            messagebox.showinfo("Sucesso", f"Arquivos processados com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro: {e}")

    root = tk.Tk()
    root.title("Processador de Arquivos")
    root.geometry("500x300")

    ttk.Label(root, text="Selecione uma Gestora:", font=("Arial", 12)).pack(pady=10)
    combo_gestoras = ttk.Combobox(root, values=[
        "Acadian", "Colchester", "Fundamenta", "Lord Abett", "Man", "Oaktree", "Pearl Diver", "Zeno"
    ], state="readonly", font=("Arial", 10))
    combo_gestoras.pack(pady=5)

    ttk.Label(root, text="Selecione o Diretório:", font=("Arial", 12)).pack(pady=10)
    entry_diretorio = ttk.Entry(root, width=50, font=("Arial", 10))
    entry_diretorio.pack(pady=5)
    ttk.Button(root, text="Selecionar Diretório", command=selecionar_diretorio).pack(pady=5)

    ttk.Button(root, text="Processar", command=processar, width=20).pack(pady=20)

    root.mainloop()

create_ui()
