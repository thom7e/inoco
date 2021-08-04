from tkinter import *
import os
import datetime
import xlsxwriter


def button_action():
    root_src_dir = eingabefeld.get() #Input
    workbook = xlsxwriter.Workbook(f"{root_src_dir}\\Planliste.xlsx")
    worksheet = workbook.add_worksheet("Inhaltsverzeichnis")
    cell_format = workbook.add_format({'bold': True})
    cell_format2 = workbook.add_format({'bold': True, 'font_size': '14'}) #Headings of the Cells
    worksheet.write(f'B1', "Bezeichnung", cell_format2)
    worksheet.write(f'C1', "Ablegedatum", cell_format2)
    n = 2
    m = 2

    with open("inhaltsverzeichnis.txt", "w") as datei:
        for src_dir, dirs, files in os.walk(root_src_dir):  # os.walkthrough
            ordnername = src_dir.split("\\")[-1]
            datei.write(f"Ordnername: {ordnername}")
            print(f"Ordnername: {ordnername}")
            worksheet.write(f'A{n}', f'{ordnername}', cell_format)
            n += 1
            m += 1
            for file_ in files:
                datum = os.path.getmtime(f"{src_dir}\\{file_}")
                v = datetime.datetime.fromtimestamp(datum)
                x = v.strftime('%Y\\%m\\%d')
                # src_file = os.path.join(src_dir, file_)
                datei.write(f"Plan \"{file_}\" vom {x}")
                print(f"Plan \"{file_}\" vom {x}")
                worksheet.write(f'B{n}', f'{file_}')
                n += 1
                worksheet.write(f'C{m}', f'{x}')
                m += 1

    workbook.close()


# Ein Fenster erstellen
fenster = Tk()
# Fenstertitel erstellen
fenster.title("InoCo")

# Fenster stellen
my_label = Label(fenster, text="Kopiere den Pfad hier rein", padx=120)

# Hier macht der Benutzer seine Eingabe
eingabefeld = Entry(fenster, bd=5, width=40)

rechnen_button = Button(fenster, text="Excel-Datei erstellen", command=button_action, pady=5)
exit_button = Button(fenster, text="Beenden", command=fenster.quit, pady=5)

# Nun fügen wir die Komponenten unserem Fenster in der gewünschten Reihenfolge hinzu
my_label.pack()
eingabefeld.pack()
rechnen_button.pack()

exit_button.pack()
fenster.bind('<Return>', button_action)
rechnen_button.focus_set()

# In der Ereignissschleife auf Eingabe des Benutzers warten
fenster.mainloop()
# baugrube()
