import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from tkcalendar import Calendar
from tkcalendar import DateEntry
from datetime import date
import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from datetime import date

#global variables
file_paths = []
save_path = 0
variant = 0
start_date = 0
end_date = 0

def open_csv_files():
    global file_paths
    paths = filedialog.askopenfilenames(
        title="CSV-Datei, bzw. Dateien auswählen",
        filetypes=[("CSV Dateien", "*.csv")]
    )
    if not paths: 
        return
    
    if len(paths) > 5:
        messagebox.showwarning("Zu viele Dateien", "Bitte wählen Sie maximal 5 .csv-Dateien aus")
        return
    file_paths = list(paths)
    
    lbl_status.config(text=f"{len(file_paths)} Datei(en) ausgewählt")

def submit_dates():
    global start_date
    global end_date
    start_date = start_cal.get_date()
    start_date = pd.to_datetime(start_date)
    end_date = end_cal.get_date()
    end_date = pd.to_datetime(end_date)
    # if start_date or end_date < pd.to_datetime("01.02.2025", format="%d.%m.%Y"):
    #     messagebox.showerror("Ungültige Auswahl", "Das ausgewählte Datum darf nicht vor dem 01.02.2025 liegen.") 
    if start_date > end_date:
        messagebox.showerror("Ungültige Auswahl", "Das Startdatum darf nicht nach dem Enddatum liegen.")
    else:
        print("Startdatum:", start_date)
        print("Enddatum:", end_date)

def select_save_path():
    global save_path
    save_path = filedialog.askdirectory()

def build_dataframe():
    list_of_df = []
    global df
    #read all csv and append to new df list
    for file in file_paths:
        df = pd.read_csv(file, sep = ";", usecols = [1, 6, 10, 11, 12],skiprows = 2, header = None)
        list_of_df.append(df)  
    #concat all dfs with all stations
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
    print(df.tail())
    
def filter_for_variant(event):
    global variant
    selected_value = filter_var.get()
    variant = selected_value
    print(f"Es wurde die Variante {variant} ausgewählt.")

def main_filter_func():
    if save_path and variant and start_date and end_date != 0:
        #Filter base df for start & enddate
        df_filtered = df[(df["Datum"] >= start_date) & (df["Datum"] <= end_date)]
        #function call
        df_grouped_detailed = detailed_dataframe(df_filtered)
        df_grouped_super_detailed = super_detailed_dataframe(df_filtered)
        if start_date == end_date:
            fig_l = create_pareto(df_grouped_detailed, "L")
            fig_r = create_pareto(df_grouped_detailed, "R")
            excel_export(df_grouped_detailed, df_grouped_super_detailed)
            pdf_report_export(fig_l, fig_r)    
            messagebox.showinfo("✅ Export erfolgreich", "Der Schraubreport sowie der Prüfbericht wurden erfolgreich exportiert.")
        else:
            excel_export(df_grouped_detailed, df_grouped_super_detailed)
            messagebox.showinfo("✅ Export erfolgreich", "Der Schraubreport wurde erfolgreich durchgeführt.")
    else:
        messagebox.showerror("Fehlerhafte Eingabe", "Achtung es wurden nicht alle Parameter gesetzt.")  
        
def detailed_dataframe(df_filtered): 
    #Group dataframe into correct shape
    df_grouped_detailed = df_filtered.groupby([df_filtered["Datum"].dt.date, "Station", "Bauteil", "Status"]).size().unstack(fill_value=0)
    #Filter for the set variant needed
    if variant == "Alle Varianten":
       pass
    else:
        df_grouped_detailed = df_grouped_detailed.loc[df_grouped_detailed.index.get_level_values("Bauteil").str.contains(f"{variant}", case=False, na=False)]
    #Calc relative failure percentage and set new df column
    df_grouped_detailed["Relativer Fehler in %"] = (df_grouped_detailed["Verschraubung NIO"] / 
                                     (df_grouped_detailed["Verschraubung IO"] + 
                                      df_grouped_detailed["Verschraubung NIO"]) * 100).round(1)
    #sort for most NIO
    df_grouped_detailed = df_grouped_detailed.reset_index()
    if "Relativer Fehler in %" in df_grouped_detailed.columns:
        df_grouped_detailed = df_grouped_detailed.sort_values(by=["Datum", "Station", "Relativer Fehler in %"], ascending=[True, True, False])
    df_grouped_detailed = df_grouped_detailed.set_index(["Datum", "Station", "Bauteil"])
    return (df_grouped_detailed)

def super_detailed_dataframe(df_filtered):
    #groupby aggregation
    df_grouped_super_detailed = df_filtered.groupby([df_filtered["Datum"].dt.date,"Station", "Bauteil", "Statusinfo"]).size().unstack()
    #delete cols without entry
    df_grouped_super_detailed = df_grouped_super_detailed.loc[:, (df_grouped_super_detailed != 0).any(axis=0)]
    #fill up int values 
    df_grouped_super_detailed = df_grouped_super_detailed.fillna(0).astype(int)
    cols = df_grouped_super_detailed.columns.tolist()
    if "Verschraubung IO" in cols: 
        cols = ["Verschraubung IO"] + [col for col in cols if col != "Verschraubung IO"] 
        df_grouped_super_detailed = df_grouped_super_detailed[cols]
    if variant == "Alle Varianten":
        return(df_grouped_super_detailed)
    else:
        df_grouped_super_detailed = df_grouped_super_detailed.loc[df_grouped_super_detailed.index.get_level_values("Bauteil").str.contains(f"{variant}", case=False, na=False)]
        return(df_grouped_super_detailed)
    
def excel_export(df_grouped_detailed, df_grouped_super_detailed):
    if start_date == end_date:
        with pd.ExcelWriter(f"{save_path}/Schraubreport_{start_date.date()}_{variant}.xlsx") as writer:
            df_grouped_detailed.to_excel(writer, sheet_name = f"{variant}, IO vs nIO")
            df_grouped_super_detailed.to_excel(writer, sheet_name = f"{variant}, detailliert")
    else:
        with pd.ExcelWriter(f"{save_path}/Schraubreport_{start_date.date()},{end_date.date()}_{variant}.xlsx") as writer:
            df_grouped_detailed.to_excel(writer, sheet_name = f"{variant} IO vs nIO")
            df_grouped_super_detailed.to_excel(writer, sheet_name = f"{variant}, detailliert")

def create_pareto(df_grouped_detailed, side):
    df_pareto = df_grouped_detailed
    df_pareto = df_pareto[df_pareto.index.get_level_values('Bauteil').str.contains(f'{variant} {side}')]
    df_pareto = df_pareto.sort_values('Verschraubung NIO', ascending=False)
    total_nio = df_pareto['Verschraubung NIO'].sum()
    df_pareto['Kumulierter Anteil in %'] = df_pareto['Verschraubung NIO'].cumsum() / total_nio * 100
    df_pareto = df_pareto.reset_index()
    df_pareto = df_pareto.drop(["Datum", "Station", "Relativer Fehler in %"], axis = 1)
    df_pareto = df_pareto[df_pareto["Verschraubung NIO"] != 0]
    df_pareto['Fehleranteil in %'] = df_pareto['Verschraubung NIO'] / total_nio * 100
    
    #Plot
    
    fig, ax1 = plt.subplots(figsize=(12, 7))

    bars = ax1.bar(df_pareto['Bauteil'], df_pareto['Verschraubung NIO'], color='royalblue')
    ax1.set_ylabel('Anzahl Verschraubung NIO', color='royalblue')
    ax1.tick_params(axis='y', labelcolor='royalblue')
    ax1.set_xticks(range(len(df_pareto['Bauteil'])))
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
    plt.title(f'Pareto-Diagramm der Fehlverschraubungen am {start_date.date()}, {variant} {side}')
    return fig

def pdf_report_export(fig_l, fig_r):
    filename = f"{save_path}/Prüfbericht_{variant}_{start_date.date()}.pdf"
    #PDF export
    with PdfPages(filename) as pdf:
        fig_l.set_size_inches(11.69, 8.27)
        fig_r.set_size_inches(11.69, 8.27)
        pdf.savefig(fig_l)
        pdf.savefig(fig_r)
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
                                command=submit_dates)
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
                            command=main_filter_func,
                            style="Export.TButton")
    btn_export.grid(row=7, column=0, pady=20, sticky="ew")

    #Separator
    ttk.Separator(root, orient="horizontal") \
        .grid(row=8, column=0, sticky="ew", pady=15)

    #Author + Version
    lbl_version = ttk.Label(root,
                            text="Phillip Kusinski, V1.0",
                            style="TLabel") 
    lbl_version.grid(row=9, column=0, sticky="e")

    root.mainloop()