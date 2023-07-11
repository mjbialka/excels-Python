from tkinter import Tk, Label, Button, OptionMenu, StringVar, filedialog
import main
import os

root = Tk()
root.title("Aplikacja do wczytywania danych z txt do excel")
root.geometry("280x200")

# Ustawienie koloru tła na ciemny
root.configure(bg="#333333")

txt_file = None
delimiter = None
destination_path = None
csv_file = None

def choose_file():
    global txt_file
    txt_file = filedialog.askopenfilename(filetypes=[("Pliki tekstowe", "*.txt")])
    # Działania wykonywane po wybraniu pliku wsadowego

label1 = Label(root, text="Lokalizacja pliku wsadowego (txt):", bg="#333333", fg="white")
label1.pack()

button1 = Button(root, text="Wybierz plik wsadowy", command=choose_file)
button1.pack()

delimiter_options = ['- (myślnik)', '; (średnik)', ', (przecinek)']
selected_delimiter = StringVar(root)
selected_delimiter.set(delimiter_options[0])

delimiter_label = Label(root, text="Znak oddzielający kolumny:", bg="#333333", fg="white")
delimiter_label.pack()

delimiter_menu = OptionMenu(root, selected_delimiter, *delimiter_options)
delimiter_menu.config(bg="#333333", fg="white")
delimiter_menu.pack()

def choose_destination():
    global destination_path
    destination_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Pliki Excel", "*.xlsx")])
    # Działania wykonywane po wybraniu pliku docelowego

label2 = Label(root, text="Lokalizacja pliku docelowego (Excel):", bg="#333333", fg="white")
label2.pack()

button2 = Button(root, text="Wybierz plik docelowy", command=choose_destination)
button2.pack()

def execute_script():
    global txt_file, delimiter, destination_path, csv_file
    delimiter_option = selected_delimiter.get()
    if delimiter_option == '- (myślnik)':
        delimiter = '-'
    elif delimiter_option == '; (średnik)':
        delimiter = ';'
    elif delimiter_option == ', (przecinek)':
        delimiter = ','
    else:
        delimiter = ''

    if txt_file and delimiter and destination_path:
        csv_file = main.generate_csv_from_txt(txt_file, delimiter)
        main.convert_csv_to_excel(csv_file, destination_path)

execute_button = Button(root, text="Wykonaj", command=execute_script)
execute_button.pack()

root.mainloop()
