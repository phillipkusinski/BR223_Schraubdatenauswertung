"""
Author: Phillip Kusinski
GUI tool for analyzing and exporting screw assembly data for BR223 production reports
"""

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
from datetime import date
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

#global variables
file_paths = []
save_path = 0
variant = 0
start_date = 0
end_date = 0
calendarweek = 0
list_of_days = ["Mo.", "Di.", "Mi.", "Do.", "Fr.", "Sa.", "So."]

#function definitions
def open_csv_files():
    global file_paths
    paths = filedialog.askopenfilenames(
        title="CSV-Datei, bzw. Dateien auswählen",
        filetypes=[("CSV Dateien", "*.csv")]
    )
    #return if no path was selected
    if not paths: 
        return
    #if more data was selected then there are screw robots a error must occur
    if len(paths) > 5:
        messagebox.showwarning("Zu viele Dateien", "Bitte wählen Sie maximal 5 .csv-Dateien aus")
        return
    
    file_paths = list(paths)
    lbl_status.config(text=f"{len(file_paths)} Datei(en) ausgewählt")

def submit_dates(list_of_days):
    global start_date
    global end_date
    global calendarweek
    #set start and end date
    start_date = start_cal.get_date()
    start_date = pd.to_datetime(start_date)
    end_date = end_cal.get_date()
    end_date = pd.to_datetime(end_date)
    #set cutoff date where the df should be cut
    cutoff_date = pd.to_datetime("2025-02-01")
    #cut off data that is too old
    if start_date < cutoff_date or end_date < cutoff_date:
        messagebox.showerror("Ungültige Auswahl", "Das ausgewählte Datum darf nicht vor dem 01.02.2025 liegen.") 
        return
    #start_date after end_date is not possible
    elif start_date > end_date:
        messagebox.showerror("Ungültige Auswahl", "Das Startdatum darf nicht nach dem Enddatum liegen.")
        return
    #if start_date and end_date is not in the same calendarweek an error must occur
    elif start_date.isocalendar().week != end_date.isocalendar().week:
        messagebox.showerror("Ungültige Auswahl", "Es wurde keine gültige Kalenderwoche ausgewählt")
        return
    #single date chosen
    elif start_date == end_date:
        messagebox.showinfo("Datumsauswahl", f"Ausgewähltes Datum: {list_of_days[start_date.weekday()]} {start_date.date()}")
    #calendarweek chosen
    else:
        calendarweek = start_date.isocalendar().week
        messagebox.showinfo("Datumsauswahl", f"Ausgewählte Datum: KW{calendarweek}, von {list_of_days[start_date.weekday()]} {start_date.date()} bis {list_of_days[end_date.weekday()]} {end_date.date()}")

def select_save_path():
    global save_path
    #get saveing directory from user input
    save_path = filedialog.askdirectory(
        title = "Ordner zur Abspeicherung der Prüfergebnisse auswählen."
    )
    #return if no save_path was selected
    if not save_path:
        return
    messagebox.showinfo("Ordnerwahl erfolgreich", "Es wurde erfolgreich ein Ordner zur Abspeicherung ausgewählt.")

def build_dataframe():
    list_of_df = []
    global df
    #read all csv and append to new df list
    for file in file_paths:
        df = pd.read_csv(file, sep = ";", usecols = [1, 6, 10, 11, 12],skiprows = 2, header = None)
        list_of_df.append(df)  
    #concat all csvs in the list_of_df into one big df
    df = pd.concat(list_of_df, ignore_index=True)

    #split cols with date and time and drop useless col with date+time
    df[["Datum", "Uhrzeit"]] = df[6].str.split(" ", expand = True)
    df = df.drop(6, axis = 1)

    #evaluate a minimum start date to drop older screw data
    df["Datum"] = pd.to_datetime(df["Datum"], format="%d.%m.%Y")
    df = df[df["Datum"] >= pd.to_datetime("01.02.2025", format="%d.%m.%Y")]

    #set column headers
    header = ["Station", "Status", "Statusinfo", "Bauteil", "Datum", "Uhrzeit"]
    df.columns = header
    messagebox.showinfo("Datenstruktur erfolgreich", "Es wurde erfolgreich die Datenstruktur aufgebaut.")
    
def filter_for_variant(event):
    global variant
    #get user input from dropdown menu from user
    selected_value = filter_var.get()
    variant = selected_value
    print(f"Es wurde die Variante {variant} ausgewählt.")

def main_filter_func(list_of_days):
    #if all data was set correctly for procedure the main_filter_func can be started
    if save_path and variant and start_date and end_date != 0:
        #Filter base df for start_date & end_date
        df_filtered = df[(df["Datum"] >= start_date) & (df["Datum"] <= end_date)]
        #function call to create dataframe structure
        df_grouped_detailed = detailed_dataframe(df_filtered)
        #check if created df is empty
        if df_grouped_detailed.empty == False:
            df_grouped_super_detailed = super_detailed_dataframe(df_filtered)
            #create pareto & excel for single date export
            if start_date == end_date:
                if variant != "Alle Varianten":
                    fig_l = create_pareto_single_date(df_grouped_detailed, "L", list_of_days) 
                    fig_r = create_pareto_single_date(df_grouped_detailed, "R", list_of_days)
                    excel_export(**{"IO vs nIO": df_grouped_detailed, "detailliert": df_grouped_super_detailed})
                    pdf_report_export(fig_l, fig_r)    
                    messagebox.showinfo("✅ Export erfolgreich", "Der Schraubreport sowie der Prüfbericht wurden erfolgreich exportiert.")
                else:
                    excel_export(**{"IO vs nIO": df_grouped_detailed, "detailliert": df_grouped_super_detailed})
                    messagebox.showinfo("✅ Export erfolgreich", "Der Schraubreport wurde erfolgreich exportiert.")
            #create pareto & excel for weekly export    
            else:
                df_grouped_detailed_weekly = detailed_dataframe_weekly_base(df_filtered)
                df_grouped_super_detailed_weekly = super_detailed_dataframe_weekly_base(df_filtered)
                if variant != "Alle Varianten":               
                    fig_l = create_pareto_weekly(df_grouped_detailed_weekly, "L", list_of_days)
                    fig_r = create_pareto_weekly(df_grouped_detailed_weekly, "R", list_of_days)
                    pdf_report_export(fig_l, fig_r)
                    excel_export(**{"IO vs nIO": df_grouped_detailed, "detailliert": df_grouped_super_detailed, "IO vs nIO Wochenfehler":df_grouped_detailed_weekly, "detaillierter Wochenfehler": df_grouped_super_detailed_weekly})
                    messagebox.showinfo("✅ Export erfolgreich", "Der Schraubreport sowie der wöchentliche Prüfbericht wurden erfolgreich exportiert.")
                else:
                    excel_export(**{"IO vs nIO": df_grouped_detailed, "detailliert": df_grouped_super_detailed, "IO vs nIO Wochenfehler":df_grouped_detailed_weekly, "detaillierter Wochenfehler": df_grouped_super_detailed_weekly})
                    messagebox.showinfo("✅ Export erfolgreich", "Der Schraubreport wurde erfolgreich exportiert.")
        else:
            messagebox.showinfo("Keine Daten vorhanden", "Es wurde ein Datum ausgewählt an dem keine Daten vorhanden sind.")
    else:
        messagebox.showerror("Fehlerhafte Eingabe", "Achtung es wurden nicht alle Parameter gesetzt.")  
        
def detailed_dataframe(df_filtered): 
    #Group dataframe into correct shape
    df_grouped_detailed = df_filtered.groupby([df_filtered["Datum"].dt.date, "Station", "Bauteil", "Status"]).size().unstack(fill_value=0)
    #Filter for the set variant needed
    if variant == "Alle Varianten":
       pass
    #Filter for chosen variant if variant != "Alle Varianten"
    else:
        df_grouped_detailed = df_grouped_detailed.loc[df_grouped_detailed.index.get_level_values("Bauteil").str.contains(f"{variant}", case=False, na=False)]
    #Calc relative failure percentage and set new df column
    if df_grouped_detailed.empty == True:
        return(df_grouped_detailed)
    else:
        df_grouped_detailed["Relativer Fehler in %"] = (df_grouped_detailed["Verschraubung NIO"] / 
                                        (df_grouped_detailed["Verschraubung IO"] + 
                                        df_grouped_detailed["Verschraubung NIO"]) * 100).round(1)
        #reset index to get full accesebility to the columns
        df_grouped_detailed = df_grouped_detailed.reset_index()
        #create new colum with "Relativer Fehler in %" and make it ascending
        if "Relativer Fehler in %" in df_grouped_detailed.columns:
            df_grouped_detailed = df_grouped_detailed.sort_values(by=["Datum", "Station", "Relativer Fehler in %"], ascending=[True, True, False])
        #set back index again
        df_grouped_detailed = df_grouped_detailed.set_index(["Datum", "Station", "Bauteil"])
        return (df_grouped_detailed)

def super_detailed_dataframe(df_filtered):
    #groupby aggregation
    df_grouped_super_detailed = df_filtered.groupby([df_filtered["Datum"].dt.date,"Station", "Bauteil", "Statusinfo"]).size().unstack()
    #delete cols without entry
    df_grouped_super_detailed = df_grouped_super_detailed.loc[:, (df_grouped_super_detailed != 0).any(axis=0)]
    #fill up int values 
    df_grouped_super_detailed = df_grouped_super_detailed.fillna(0).astype(int)
    #write all columns into one list
    cols = df_grouped_super_detailed.columns.tolist()
    #write "Verschraubung IO" at first place if it is available
    if "Verschraubung IO" in cols: 
        cols = ["Verschraubung IO"] + [col for col in cols if col != "Verschraubung IO"] 
        df_grouped_super_detailed = df_grouped_super_detailed[cols]
    if variant == "Alle Varianten":
        return(df_grouped_super_detailed)
    else:
        df_grouped_super_detailed = df_grouped_super_detailed.loc[df_grouped_super_detailed.index.get_level_values("Bauteil").str.contains(f"{variant}", case=False, na=False)]
        return(df_grouped_super_detailed)

def detailed_dataframe_weekly_base(df_filtered):
    df_grouped_detailed_weekly = df_filtered.groupby([ "Station", "Bauteil", "Status"]).size().unstack(fill_value=0)
    #Filter for the set variant needed
    if variant == "Alle Varianten":
       pass
    else:
        df_grouped_detailed_weekly = df_grouped_detailed_weekly.loc[df_grouped_detailed_weekly.index.get_level_values("Bauteil").str.contains(f"{variant}", case=False, na=False)]
    #Calc relative failure percentage and set new df column
    df_grouped_detailed_weekly["Relativer Fehler in %"] = (df_grouped_detailed_weekly["Verschraubung NIO"] / 
                                     (df_grouped_detailed_weekly["Verschraubung IO"] + 
                                      df_grouped_detailed_weekly["Verschraubung NIO"]) * 100).round(1)
    #sort for most NIO
    df_grouped_detailed_weekly = df_grouped_detailed_weekly.reset_index()
    if "Relativer Fehler in %" in df_grouped_detailed_weekly.columns:
        df_grouped_detailed_weekly = df_grouped_detailed_weekly.sort_values(by=["Station", "Relativer Fehler in %"], ascending=[True, False])
    df_grouped_detailed_weekly = df_grouped_detailed_weekly.set_index(["Station", "Bauteil"])
    return (df_grouped_detailed_weekly)

def super_detailed_dataframe_weekly_base(df_filtered):
    #groupby aggregation
    df_grouped_super_detailed_weekly = df_filtered.groupby(["Station", "Bauteil", "Statusinfo"]).size().unstack()
    #delete cols without entry
    df_grouped_super_detailed_weekly = df_grouped_super_detailed_weekly.loc[:, (df_grouped_super_detailed_weekly != 0).any(axis=0)]

    #fill up int values 
    df_grouped_super_detailed_weekly = df_grouped_super_detailed_weekly.fillna(0).astype(int)

    cols = df_grouped_super_detailed_weekly.columns.tolist()
    if "Verschraubung IO" in cols: 
        cols = ["Verschraubung IO"] + [col for col in cols if col != "Verschraubung IO"] 
        df_grouped_super_detailed_weekly = df_grouped_super_detailed_weekly[cols]
    if variant == "Alle Varianten":
        return(df_grouped_super_detailed_weekly)   
    else:
        df_grouped_super_detailed_weekly = df_grouped_super_detailed_weekly.loc[df_grouped_super_detailed_weekly.index.get_level_values("Bauteil").str.contains(f"{variant}", case=False, na=False)]
        return(df_grouped_super_detailed_weekly)
    
def excel_export(**kwargs):
    #set excel filename for single day and for weekly output
    if start_date == end_date:
        filename = f"{save_path}/Schraubreport_{variant}_{start_date.date()}.xlsx"
    else:
        filename = f"{save_path}/Schraubreport_{variant}_KW{calendarweek}_{start_date.date()},{end_date.date()}.xlsx"   
    #export excel function with variable **kwargs input     
    with pd.ExcelWriter(filename) as writer:
        for name, df in kwargs.items():
            sheet_name = f"{name}"
            df.to_excel(writer, sheet_name=sheet_name)

def create_pareto_single_date(df_grouped_detailed, side, list_of_days):
    #filter for variant
    df_pareto = df_grouped_detailed[df_grouped_detailed.index.get_level_values('Bauteil').str.contains(f'{variant} {side}')]
    #sort "Verschraubung NIO" ascending for pareto preperation
    df_pareto = df_pareto.sort_values('Verschraubung NIO', ascending=False)
    #calc total_nio of col 
    total_nio = df_pareto['Verschraubung NIO'].sum()
    #calc "Kumulierter Anteil in %"
    df_pareto['Kumulierter Anteil in %'] = df_pareto['Verschraubung NIO'].cumsum() / total_nio * 100
    #reset columns index to get full access
    df_pareto = df_pareto.reset_index()
    #drop useless cols for pareto
    df_pareto = df_pareto.drop(["Datum", "Station", "Relativer Fehler in %"], axis = 1)
    #df should only contain entry where "Verschraubung NIO" is not 0
    df_pareto = df_pareto[df_pareto["Verschraubung NIO"] != 0]
    #Calc realtive failure percentages
    df_pareto['Fehleranteil in %'] = df_pareto['Verschraubung NIO'] / total_nio * 100
    
    #Plot
    
    fig, ax1 = plt.subplots(figsize=(12, 7))

    x = range(len(df_pareto))
    bars = ax1.bar(x, df_pareto['Verschraubung NIO'], color='royalblue')
    ax1.set_ylabel('Anzahl Verschraubung NIO', color='royalblue')
    ax1.tick_params(axis='y', labelcolor='royalblue')
    ax1.set_xticks(x)
    ax1.set_xticklabels(df_pareto['Bauteil'], rotation=90, ha='right')

    for bar, percent in zip(bars, df_pareto['Fehleranteil in %']):
        height = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2, height + 0.2, f'{percent:.1f}%', ha='center', va='bottom', fontsize=7)

    ax2 = ax1.twinx()
    ax2.plot(df_pareto['Bauteil'], df_pareto['Kumulierter Anteil in %'], color='firebrick', marker='o', linewidth=2)
    ax2.set_ylabel('Kumulierte Fehleranteile in %', color='firebrick')
    ax2.tick_params(axis='y', labelcolor='firebrick')
    ax2.axhline(80, color='gray', linestyle='--', label='80 % Schwelle')

    plt.tight_layout(pad=1.5)
    plt.title(f'Pareto-Diagramm der Fehlverschraubungen am {list_of_days[start_date.weekday()]} {start_date.date()}, {variant} {side}')
    return fig

def create_pareto_weekly(df_grouped_detailed_weekly, side, list_of_days):
    #filter for variant
    df_pareto = df_grouped_detailed_weekly[df_grouped_detailed_weekly.index.get_level_values('Bauteil').str.contains(f'{variant} {side}')]
    #reset index for deeper level
    df_pareto = df_pareto.reset_index(level = ['Bauteil'])
    df_pareto = df_pareto[['Bauteil', 'Verschraubung IO', 'Verschraubung NIO']]
    df_pareto = df_pareto.reset_index()
    #drop useless col
    df_pareto = df_pareto.drop(['Station'], axis = 1)
    #internal sum over col
    df_pareto = df_pareto.groupby('Bauteil', as_index=False).sum()
    #sort "Verschraubung NIO" ascending for pareto preperation
    df_pareto = df_pareto.sort_values('Verschraubung NIO', ascending=False)
    #calc total_nio of col 
    total_nio = df_pareto['Verschraubung NIO'].sum()
    #calc "Kumulierter Anteil in %"
    df_pareto['Kumulierter Anteil in %'] = df_pareto['Verschraubung NIO'].cumsum() / total_nio * 100
    #df should only contain entry where "Verschraubung NIO" is not 0
    df_pareto = df_pareto[df_pareto["Verschraubung NIO"] != 0]
    #Calc realtive failure percentages
    df_pareto['Fehleranteil in %'] = df_pareto['Verschraubung NIO'] / total_nio * 100

    #PLOT

    fig, ax1 = plt.subplots(figsize=(12, 6))

    x = range(len(df_pareto))
    bars = ax1.bar(x, df_pareto['Verschraubung NIO'], color='royalblue')
    ax1.set_ylabel('Anzahl Verschraubung NIO', color='royalblue')
    ax1.tick_params(axis='y', labelcolor='royalblue')
    ax1.set_xticks(x)
    ax1.set_xticklabels(df_pareto['Bauteil'], rotation=90, ha='right')

    for bar, percent in zip(bars, df_pareto['Fehleranteil in %']):
        height = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2, height + 0.2, f'{percent:.1f}%', ha='center', va='bottom', fontsize=7)

    ax2 = ax1.twinx()
    ax2.plot(df_pareto['Bauteil'], df_pareto['Kumulierter Anteil in %'], color='firebrick', marker='o', linewidth=2)
    ax2.set_ylabel('Kumulierte Fehleranteile in %', color='firebrick')
    ax2.tick_params(axis='y', labelcolor='firebrick')
    ax2.axhline(80, color='gray', linestyle='--', label='80 % Schwelle')
    
    plt.tight_layout(pad=1.5)
    plt.title(f'Pareto-Diagramm der Fehlverschraubungen KW{calendarweek}, {list_of_days[start_date.weekday()]} {start_date.date()} bis {list_of_days[end_date.weekday()]} {end_date.date()}, {variant} {side}')
    return fig

def pdf_report_export(fig_l, fig_r):
    #day output
    if start_date == end_date:
        filename = f"{save_path}/Prüfbericht_{variant}_{start_date.date()}.pdf"
        #PDF export
        with PdfPages(filename) as pdf:
            #set export size to DINA4
            fig_l.set_size_inches(11.69, 8.27)
            fig_r.set_size_inches(11.69, 8.27)
            pdf.savefig(fig_l)
            pdf.savefig(fig_r)
            #set metadata
            pdf.infodict()['Title'] = 'BR223 Prüfbericht Schraubzelle'
            pdf.infodict()['Author'] = 'Phillip Kusinski'
            pdf.infodict()['Subject'] = 'Dieser Prüfbericht wurde mit der Software Schraubdatenauswertung_BR223 erstellt.'
            pdf.infodict()['Keywords'] = 'BR223, Screwing, Report, Pareto'
    #weekly output
    else:
        filename = f"{save_path}/Prüfbericht_{variant}_KW{calendarweek}_{start_date.date()}-{end_date.date()}.pdf"
        #PDF export
        with PdfPages(filename) as pdf:
            
            fig_l.set_size_inches(11.69, 8.27)
            fig_r.set_size_inches(11.69, 8.27)
            pdf.savefig(fig_l)
            pdf.savefig(fig_r)
            #set metadata
            pdf.infodict()['Title'] = 'BR223 Prüfbericht Schraubzelle'
            pdf.infodict()['Author'] = 'Phillip Kusinski'
            pdf.infodict()['Subject'] = 'Dieser Prüfbericht wurde mit der Software Schraubdatenauswertung_BR223 erstellt.'
            pdf.infodict()['Keywords'] = 'BR223, Screwing, Report, Pareto'
    print(f"PDF erfolgreich exportiert: {filename}")    

if __name__ == "__main__":  
    #Setup Main Window
    root = tk.Tk()
    root.title("BR223 Schraubauswertung")
    root.resizable(False, False)
    #iconbitmap does not work with .exe build without bigger changes
    #root.iconbitmap("ressources/logo_yf.ico")

    #global Padding
    root.configure(padx=20, pady=20, bg="#f0f0f0")

    #style config
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("TFrame", background="#f0f0f0")
    style.configure("TLabel", background="#f0f0f0")
    style.configure("Export.TButton",
                    font=("Arial", 16, "bold"),
                    foreground="white",
                    background="#28a745")
    style.map("Export.TButton",
            background=[("active", "#1e7e34")],
            foreground=[("active", "white")])

    #CSV import
    frame_csv = ttk.Frame(root)
    frame_csv.grid(row=0, column=0, sticky="ew")
    root.columnconfigure(0, weight=1)
    frame_csv.columnconfigure(0, weight=1)
    frame_csv.columnconfigure(1, weight=1)

    btn_load_csv = ttk.Button(frame_csv,
                            text="📂 CSV-Datei öffnen",
                            command=open_csv_files)
    btn_load_csv.grid(row=0, column=0, sticky="ew")

    lbl_status = ttk.Label(frame_csv,
                        text="0 Dateien ausgewählt")
    lbl_status.grid(row=0, column=1, sticky="w", padx=(20, 0))

    btn_submit_csv = ttk.Button(
        frame_csv,
        text="Erstelle Datenstruktur",
        command=build_dataframe
    )
    btn_submit_csv.grid(row=1, column=0, columnspan = 2, sticky="ew", pady=10)

    #Separator
    ttk.Separator(root, orient="horizontal") \
        .grid(row=1, column=0, sticky="ew", pady=15)

    #Choose Date
    frame_dates = ttk.Frame(root)
    frame_dates.grid(row=2, column=0, sticky="ew")
    frame_dates.columnconfigure((0, 1), weight=1, uniform="a")

    ttk.Label(frame_dates, text="Startdatum auswählen:") \
        .grid(row=0, column=0, sticky="w")
    start_cal = DateEntry(frame_dates,
                        width=18,
                        background="darkblue",
                        foreground="white",
                        borderwidth=2,
                        maxdate=date.today())
    start_cal.grid(row=1, column=0, sticky="w", pady=5)

    ttk.Label(frame_dates, text="Enddatum auswählen:") \
        .grid(row=0, column=1, sticky="w", padx=(20, 0))
    end_cal = DateEntry(frame_dates,
                        width=18,
                        background="darkblue",
                        foreground="white",
                        borderwidth=2,
                        maxdate=date.today())
    end_cal.grid(row=1, column=1, sticky="w", padx=(20, 0), pady=5)

    btn_submit_dates = ttk.Button(frame_dates,
                                text="✅ Datierung bestätigen",
                                command=lambda: submit_dates(list_of_days))
    btn_submit_dates.grid(row=2, column=0, columnspan=2,
                        sticky="ew", pady=10)

    ttk.Separator(root, orient="horizontal") \
        .grid(row=3, column=0, sticky="ew", pady=15)

    #Filter Variant
    frame_variant = ttk.Frame(root)
    frame_variant.grid(row=4, column=0, sticky="ew")

    ttk.Label(frame_variant, text="Varianten auswählen:") \
        .grid(row=0, column=0, sticky="w")
    filter_var = tk.StringVar(value="Alle Varianten")
    combo = ttk.Combobox(frame_variant,
                        textvariable=filter_var,
                        values=[
                            "Alle Varianten",
                            "FAT",
                            "FOT-V",
                            "FOT-W",
                            "FOT-Z"
                        ],
                        state="readonly")
    combo.grid(row=1, column=0, sticky="ew", pady=5)
    combo.current(0)
    combo.bind("<<ComboboxSelected>>", filter_for_variant)

    #Separator
    ttk.Separator(root, orient="horizontal") \
        .grid(row=5, column=0, sticky="ew", pady=15)

    btn_select_path = ttk.Button(root,
                            text="📂 Speicherpfad auswählen",
                            command=select_save_path)
    btn_select_path.grid(row=6, column=0, sticky="ew")

    #Export
    btn_export = ttk.Button(root,
                            text="Export starten",
                            command=lambda: main_filter_func(list_of_days),
                            style="Export.TButton")
    btn_export.grid(row=7, column=0, pady=20, sticky="ew")

    #Separator
    ttk.Separator(root, orient="horizontal") \
        .grid(row=8, column=0, sticky="ew", pady=15)

    #Author + Version
    lbl_version = ttk.Label(root,
                            text="Phillip Kusinski, V1.1",
                            style="TLabel") 
    lbl_version.grid(row=9, column=0, sticky="e")

    root.mainloop()