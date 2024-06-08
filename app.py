import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends.backend_pdf import PdfPages
import warnings

def load_file():
    filepath = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )
    if not filepath:
        return
    try:
        global df
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            df = pd.read_excel(filepath, engine='openpyxl')
            if any("Conditional Formatting extension is not supported" in str(warning.message) for warning in w):
                messagebox.showwarning("Warning", "Conditional Formatting not supported and has been removed.")
        if df.empty:
            raise ValueError("O arquivo Excel está vazio.")
        display_data(df)
    except ValueError as ve:
        messagebox.showerror("Error", str(ve))
    except FileNotFoundError:
        messagebox.showerror("Error", "Arquivo não encontrado.")
    except pd.errors.ExcelFileError as efe:
        messagebox.showerror("Error", f"Erro ao abrir o arquivo Excel: {str(efe)}")
    except Exception as e:
        messagebox.showerror("Error", f"Falha ao ler o arquivo\n{str(e)}")

def display_data(dataframe):
    for widget in data_frame.winfo_children():
        widget.destroy()
    
    tree = ttk.Treeview(data_frame)
    tree.pack(fill=tk.BOTH, expand=True)
    
    tree["column"] = list(dataframe.columns)
    tree["show"] = "headings"
    
    for col in tree["columns"]:
        tree.heading(col, text=col)
        tree.column(col, anchor=tk.CENTER)
    
    for index, row in dataframe.iterrows():
        tree.insert("", "end", values=list(row))
    
    global selected_columns
    selected_columns = tk.StringVar(value=list(dataframe.columns))
    global columns_listbox
    columns_listbox = tk.Listbox(selection_frame, listvariable=selected_columns, selectmode=tk.MULTIPLE)
    columns_listbox.pack(fill=tk.BOTH, expand=True)

def plot_data():
    selected_indices = columns_listbox.curselection()
    if not selected_indices:
        messagebox.showerror("Error", "Nenhuma coluna selecionada para plotar.")
        return
    
    selected_cols = [columns_listbox.get(i) for i in selected_indices]
    fig, ax = plt.subplots(figsize=(6, 4))
    df[selected_cols].plot(kind='bar', ax=ax)
    
    global canvas
    if canvas:
        canvas.get_tk_widget().pack_forget()
    
    canvas = FigureCanvasTkAgg(fig, master=plot_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

def convert_to_pdf():
    if not canvas:
        messagebox.showerror("Error", "Nenhum dado para converter em PDF.")
        return
    
    filepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if not filepath:
        return
    
    with PdfPages(filepath) as pdf:
        pdf.savefig(canvas.figure)

window = tk.Tk()
window.title("Visualizador de Dados do Excel")

frame = tk.Frame(window)
frame.pack(pady=20)

btn_load = tk.Button(frame, text="Carregar Arquivo Excel", command=load_file)
btn_load.pack(side=tk.LEFT)

btn_plot = tk.Button(frame, text="Plotar Dados", command=plot_data)
btn_plot.pack(side=tk.LEFT)

btn_convert_to_pdf = tk.Button(window, text="Converter para PDF", command=convert_to_pdf)
btn_convert_to_pdf.pack(side=tk.BOTTOM, pady=10)

data_frame = tk.Frame(window)
data_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

selection_frame = tk.Frame(window)
selection_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

plot_frame = tk.Frame(window)
plot_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

footer_label = tk.Label(window, text="Powered by Leonardo D'avila", foreground="gray")
footer_label.pack(side=tk.BOTTOM)

df = None
canvas = None
columns_listbox = None

window.mainloop()
