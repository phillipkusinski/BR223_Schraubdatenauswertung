import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import Calendar
from tkcalendar import DateEntry
from datetime import date
import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from datetime import date

filepath = 0
def open_csv_file():
    global filepath
    filepath = filedialog.askopenfilename(
        title="CSV-Datei auswählen",
        filetypes=[("CSV Dateien", "*.csv")]
    )
    if filepath:
        print("Ausgewählte Datei:", filepath)
        
def submit_dates():
    start_date = start_cal.get_date()
    start_date = pd.to_datetime(start_date)
    end_date = end_cal.get_date()
    end_date = pd.to_datetime(end_date)
    if start_date > end_date:
        messagebox.showerror("Ungültige Auswahl", "Das Startdatum darf nicht nach dem Enddatum liegen.")
    elif filepath == 0:
        messagebox.showerror("Ungültige Auswahl", "Es wurde keine .csv Datei ausgewählt.")
    else:
        print("Startdatum:", start_date)
        print("Enddatum:", end_date)
        dataframe_manipulation(filepath, start_date, end_date)
 
def dataframe_manipulation(filepath, start_date, end_date):
    #load raw data
    df = pd.read_csv(filepath, sep = ";", usecols = [1, 6, 10, 12],skiprows = 2, header = None)
    #split cols with date and time and drop useless col with date+time
    df[["Datum", "Uhrzeit"]] = df[6].str.split(" ", expand = True)
    df = df.drop(6, axis = 1)
    #evaluate a minimum start date to drop older screw data
    df["Datum"] = pd.to_datetime(df["Datum"], format="%d.%m.%Y")
    df = df[df["Datum"] >= pd.to_datetime("01.07.2025", format="%d.%m.%Y")]
    #set column headers
    header = ["Station", "Status", "Bauteil", "Datum", "Uhrzeit"]
    df.columns = header
    station = df.loc[0, "Station"]
    #filter minimum date
    df_filtered = df[(df["Datum"] >= start_date) & (df["Datum"] <= end_date)]
    
    #Export detailed screw_data
    df_grouped_detailed = df_filtered.groupby([df_filtered["Datum"].dt.date, "Bauteil", "Status"]).size().unstack(fill_value=0)
    
    #Filtering for plot
    component = ["FAT L", "FAT R", "FOT-V L", "FOT-V R", "FOT-W L", "FOT-W R", "FOT-Z L", "FOT-Z R"]
    df_filtered = df_filtered[df_filtered["Bauteil"].str.contains("|".join(component), na=False)].copy()
    df_filtered["Bauteil"] = df_filtered["Bauteil"].apply(lambda x: next((b for b in component if b in x), x))
    df_grouped_plot = df_filtered.groupby([df_filtered["Datum"].dt.date, "Bauteil", "Status"]).size().unstack(fill_value=0)
    
    #Create plot
    df_grouped_plot = df_grouped_plot.reset_index()
    if "Verschraubung NIO" in df_grouped_plot.columns:
        pass
    else:
        df_grouped_plot["Verschraubung NIO"] = 0
    if "Verschraubung IO" in df_grouped_plot.columns:
        pass
    else:
        df_grouped_plot["Verschraubung IO"] = 0
    df_grouped_plot_copy = df_grouped_plot.copy()
    df_grouped_plot_copy["Prozentualer Fehler"] = (df_grouped_plot_copy['Verschraubung NIO'] / (df_grouped_plot_copy['Verschraubung IO'] + df_grouped_plot_copy['Verschraubung NIO'])) * 100
    #Filterung nach FAT
    df_grouped_plot_fat = df_grouped_plot_copy[df_grouped_plot_copy["Bauteil"].str.contains("FAT", na=False)]
    #Filterung nach FOT V
    df_grouped_plot_fotv = df_grouped_plot_copy[df_grouped_plot_copy["Bauteil"].str.contains("FOT-V", na=False)]
    #Filterung nach FOT W
    df_grouped_plot_fotw= df_grouped_plot_copy[df_grouped_plot_copy["Bauteil"].str.contains("FOT-W", na=False)]
    #Filterung nach FOT Z
    df_grouped_plot_fotz = df_grouped_plot_copy[df_grouped_plot_copy["Bauteil"].str.contains("FOT-Z", na=False)]
    
    #iteration list 
    plots = [(df_grouped_plot_fat, "FAT"),(df_grouped_plot_fotv, "FOT-V"),(df_grouped_plot_fotw, "FOT-W"),(df_grouped_plot_fotz, "FOT-Z")]

    #create pdf report
    with PdfPages(f'Pruefbericht_Gesamt_{station}_{start_date.date()}_{end_date.date()}.pdf') as pdf:
        for df_grouped, name in plots:
            df_plot_io_nio = df_grouped[['Verschraubung IO', 'Verschraubung NIO']]
            df_plot_fehler = df_grouped[['Prozentualer Fehler']]
            has_io_nio = not df_plot_io_nio.dropna(how='all').empty and df_plot_io_nio.sum().sum() > 0
            has_fehler = not df_plot_fehler.dropna(how='all').empty and df_plot_fehler.sum().sum() > 0
            num_plots = sum([has_io_nio, has_fehler])

            if num_plots == 0:
                fig, ax = plt.subplots(figsize=(12, 6))
                ax.axis('off')
                ax.text(0.5, 0.5,
                        f'Keine Daten für {name} vorhanden\nZeitraum: {start_date.date()} - {end_date.date()}',
                        fontsize=16, fontweight='bold', ha='center', va='center')
                pdf.savefig(fig)
                plt.close(fig)
                continue

            fig, axes = plt.subplots(num_plots, 1, figsize=(12, 6 * num_plots), sharex=True)
            if num_plots == 1:
                axes = [axes]
            plot_idx = 0

            if has_io_nio:
                df_grouped.set_index(['Datum', 'Bauteil'])[['Verschraubung IO', 'Verschraubung NIO']].plot(
                    kind='bar', stacked=False, color=['darkgreen', 'red'], edgecolor='black', ax=axes[plot_idx]
                )
                axes[plot_idx].set_title(f'Verschraubungen IO vs NIO {name}, {station}\nZeitraum {start_date.date()} - {end_date.date()}',
                                        fontsize=12, fontweight='bold')
                axes[plot_idx].set_ylabel('Anzahl')
                axes[plot_idx].grid(color='gray', linewidth=0.25)
                plot_idx += 1

            if has_fehler:
                df_grouped.set_index(['Datum', 'Bauteil'])[['Prozentualer Fehler']].plot(
                    kind='bar', stacked=False, color='royalblue', edgecolor='black', ax=axes[plot_idx]
                )
                axes[plot_idx].set_title(f'Prozentualer Fehler {name}, {station}\nZeitraum {start_date.date()} - {end_date.date()}',
                                        fontsize=12, fontweight='bold')
                axes[plot_idx].set_ylabel('Fehler in %')
                axes[plot_idx].set_ylim(0, 100)
                axes[plot_idx].grid(color='gray', linewidth=0.25)

            plt.xticks(rotation=90)
            plt.tight_layout()
            #save different figs on new site
            pdf.savefig(fig) 
            plt.close(fig)

        #set PDF metadata
        pdf.infodict()['Title'] = 'Gesamter Prüfbericht IO/NIO inkl. prozentualer Fehler'
        pdf.infodict()['Author'] = 'Phillip Kusinski'
        pdf.infodict()['Subject'] = 'Analyse der Verschraubungen aller Stationen'
        
    #Export detailed report   
    df_grouped_detailed.to_excel(f'Schraubübersicht_detailliert_{station}_{start_date.date()}_{end_date.date()}.xlsx', sheet_name = f'{station}_{start_date.date()}_{end_date.date()}')
    #Messagebox for finished export
    print('Exportieren erfolgreich')
    messagebox.showinfo(message = f'Export erfolgreich')   

        
#Mainwindow
root = tk.Tk()
root.title("BR223 Schraubauswertung")
root.geometry("400x300")
root.resizable(False, False)
#photo = tk.PhotoImage(file = "Schraubdatenauswertung/logo.png")
#root.wm_iconphoto(False, photo)

#Open CSV file
tk.Button(root, text="📂 CSV-Datei öffnen", command=open_csv_file).pack(pady=10)

#Startdate
tk.Label(root, text="Startdatum auswählen:").pack()
start_cal = DateEntry(root, width=20, background='darkblue', foreground='white', borderwidth=2, maxdate = date.today())
start_cal.pack(pady=5)

#Enddate
tk.Label(root, text="Enddatum auswählen:").pack()
end_cal = DateEntry(root, width=20, background='darkblue', foreground='white', borderwidth=2, maxdate = date.today())
end_cal.pack(pady=5)

#Submit
tk.Button(root, text="✅ Auswahl bestätigen", command=submit_dates).pack(pady=15)

tk.Label(root, text="Autor: Phillip Kusinski").pack()
root.mainloop()