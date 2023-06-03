import tkinter as tk
from tkinter import filedialog
import urlStatusChecker
import threading
import configparser


def validate_entries():
    start_from_row = start_from_row_entry.get().strip()
    sheet_number = sheet_number_entry.get().strip()
    file_to_read = path_file_to_read_entry.get().strip()
    excel_path_to_write = path_excel_to_write_entry.get().strip()

    if start_from_row and sheet_number and file_to_read and excel_path_to_write:
        execute_button.config(state=tk.NORMAL)
    else:
        execute_button.config(state=tk.DISABLED)

def load_configurations():
    config = configparser.ConfigParser()
    config.read('config.ini')

    if 'CONFIGURATIONS' in config:
        configurations = config['CONFIGURATIONS']

        # Leggi i valori dal file di configurazione, gestendo i casi in cui i campi non esistono
        start_from_row = configurations.get('start_from_row', '')
        sheet_number = configurations.get('sheet_number', '')
        do_only_one_sheet = configurations.getboolean('do_only_one_sheet', False)
        file_to_read = configurations.get('file_to_read', '')
        excel_path_to_write = configurations.get('excel_path_to_write', '')

        # Inserisci i valori nei rispettivi widget, se esistono
        start_from_row_entry.delete(0, tk.END)
        start_from_row_entry.insert(tk.END, start_from_row)
        sheet_number_entry.delete(0, tk.END)
        sheet_number_entry.insert(tk.END, sheet_number)
        do_only_one_sheet_var.set(do_only_one_sheet)
        path_file_to_read_entry.delete(0, tk.END)
        path_file_to_read_entry.insert(tk.END, file_to_read)
        path_excel_to_write_entry.delete(0, tk.END)
        path_excel_to_write_entry.insert(tk.END, excel_path_to_write)

        validate_entries()  # Aggiorna lo stato del pulsante "Esegui"


def save_configurations():
    config = configparser.ConfigParser()
    config['CONFIGURATIONS'] = {
        'start_from_row': start_from_row_entry.get(),
        'sheet_number': sheet_number_entry.get(),
        'do_only_one_sheet': str(do_only_one_sheet_var.get()),
        'file_to_read': path_file_to_read_entry.get(),
        'excel_path_to_write': path_excel_to_write_entry.get()
    }

    with open('config.ini', 'w') as config_file:
        config.write(config_file)


def add_log_message(message):
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)


# Funzione per aggiungere un messaggio di errore al log_text
def add_error_message(error_message):
    log_text.insert(tk.END, f"ERROR: {error_message}\n")
    log_text.see(tk.END)


# Funzione per eseguire lo script quando viene premuto il pulsante "Esegui"
def execute_script():
    try:
        # Visualizza il messaggio "Loading..."
        log_text.insert(tk.END, "Loading...\nPlease wait. It will require some seconds,\nit depends on .xlsx file length.")
        log_text.update()

        # Recupera i valori dai campi di input dell'interfaccia grafica
        start_from_row = int(start_from_row_entry.get())
        sheet_number = int(sheet_number_entry.get())
        do_only_one_sheet = do_only_one_sheet_var.get()
        file_to_read = path_file_to_read_entry.get()
        excel_path_to_write = path_excel_to_write_entry.get() + '\\result.xlsx'

        # Esegue lo script principale
        threading.Thread(target=urlStatusChecker.execute, args=(
        start_from_row, sheet_number, do_only_one_sheet, file_to_read, excel_path_to_write, log_text)).start()

        # Salva le configurazioni
        save_configurations()

        # Posiziona lo scroll all'ultima riga
        log_text.see(tk.END)
    except Exception as e:
        add_error_message("An exception has occurred. \nCheck configurations and try again." + str(e))


# Creazione della finestra dell'interfaccia grafica
window = tk.Tk()
window.title("URL Status Checker")
window.geometry("600x800")

# Configurazione dei colori e stili
bg_color = "#757575"
element_color = "#616161"
text_color = "#ffffff"
button_color = "#616161"
button_text_color = "#ffffff"
font = ("Roboto", 14)

# Applica lo stile alla finestra
window.configure(bg=bg_color)

# Funzione per applicare uno stile uniforme ai widget
def apply_style(widget):
    widget.configure(bg=element_color, fg=text_color, font=font)


# Funzione per consentire solo l'inserimento di numeri
def validate_numeric_input(text):
    if text.isdigit() or text == "":
        return True
    else:
        return False


# Etichetta e campo di input per start_from_row
start_from_row_label = tk.Label(window, text="Start from row:")
apply_style(start_from_row_label)
start_from_row_label.pack(pady=(16, 16), padx=8)

start_from_row_entry = tk.Entry(window)
apply_style(start_from_row_entry)
start_from_row_entry.pack()

# Registra la validazione dell'input come numerico per il campo start_from_row_entry
validate_numeric_input_cmd = window.register(validate_numeric_input)
start_from_row_entry.config(validate="key", validatecommand=(validate_numeric_input_cmd, "%P"))

# Etichetta e campo di input per sheet_number
sheet_number_label = tk.Label(window, text="Sheet number:")
apply_style(sheet_number_label)
sheet_number_label.pack(pady=(16, 16), padx=8)

sheet_number_entry = tk.Entry(window)
apply_style(sheet_number_entry)
sheet_number_entry.pack()

# Registra la validazione dell'input come numerico per il campo sheet_number_entry
sheet_number_entry.config(validate="key", validatecommand=(validate_numeric_input_cmd, "%P"))

# Checkbox per do_only_one_sheet
do_only_one_sheet_var = tk.BooleanVar()
do_only_one_sheet_check = tk.Checkbutton(window, text="Selected sheet only", variable=do_only_one_sheet_var)
apply_style(do_only_one_sheet_check)
do_only_one_sheet_check.config(bg="#424242", selectcolor="#424242")
do_only_one_sheet_check.pack()


# Pulsante e campo di input per il file da leggere
def browse_file_to_read():
    file_to_read = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    path_file_to_read_entry.delete(0, tk.END)
    path_file_to_read_entry.insert(tk.END, file_to_read)


# Etichetta e campo di input per il file da leggere
file_to_read_label = tk.Label(window, text="Select source file:")
apply_style(file_to_read_label)
file_to_read_label.pack(pady=(16, 16), padx=8)

path_file_to_read_entry = tk.Entry(window)
apply_style(path_file_to_read_entry)
path_file_to_read_entry.pack()

file_to_read_button = tk.Button(window, text="Browse", command=browse_file_to_read, bg=button_color,
                                fg=button_text_color, width=10)
file_to_read_button.pack(pady=8)


# Pulsante e campo di input per il percorso di scrittura del file Excel
def browse_excel_path_to_write():
    excel_path_to_write = filedialog.askdirectory()
    path_excel_to_write_entry.delete(0, tk.END)
    path_excel_to_write_entry.insert(tk.END, excel_path_to_write)


# Etichetta e campo di input per il percorso di scrittura del file Excel
excel_path_to_write_label = tk.Label(window, text="Select destination folder:")
apply_style(excel_path_to_write_label)
excel_path_to_write_label.pack(pady=(16, 16), padx=8)

path_excel_to_write_entry = tk.Entry(window)
apply_style(path_excel_to_write_entry)
path_excel_to_write_entry.pack()

excel_path_to_write_button = tk.Button(window, text="Browse", command=browse_excel_path_to_write, bg=button_color,
                                       fg=button_text_color)
excel_path_to_write_button.pack(pady=8)

# Campo di log
log_frame = tk.Frame(window, bg=element_color)
log_frame.pack(pady=16, padx=32)

# Creazione del widget Text per il log
log_text = tk.Text(log_frame, height=10, wrap='none')
log_text.configure(fg="white")
apply_style(log_text)
log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Imposta una larghezza massima per il widget log_text
log_text.config(width=35)  # Sostituisci 80 con la larghezza massima desiderata

# Configurazione della scrollbar del log_text
scrollbar = tk.Scrollbar(log_frame, command=log_text.yview, bg=element_color)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
log_text.config(yscrollcommand=scrollbar.set)

# Pulsante "Esegui"
execute_button = tk.Button(window, text="START", command=execute_script, bg=button_color, fg=button_text_color)
execute_button.pack(pady=16)

# Carica le configurazioni salvate
load_configurations()

# Associa la funzione validate_entries al cambiamento dei campi di input
start_from_row_entry.bind('<KeyRelease>', lambda e: validate_entries())
sheet_number_entry.bind('<KeyRelease>', lambda e: validate_entries())
path_file_to_read_entry.bind('<KeyRelease>', lambda e: validate_entries())
path_excel_to_write_entry.bind('<KeyRelease>', lambda e: validate_entries())

# Chiama la funzione validate_entries all'inizio per inizializzare lo stato del pulsante "Esegui"
validate_entries()

# Avvia l'interfaccia grafica
window.mainloop()
