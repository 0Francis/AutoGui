# monthly_etl_gui.py
# -*- coding: utf-8 -*-

import os
import sys
import traceback
import pandas as pd
import numpy as np
from tkinter import Tk, Label, Entry, Button, StringVar, filedialog, messagebox, Frame, LEFT, RIGHT, X
from tkinter import ttk

# ---------- YOUR EXISTING BUSINESS LOGIC (from automation file.txt), lightly wrapped ----------

with_customer_issues = [
    "Customer : Device Powered off",
    "Customer : Device powered off",
    "Customer : Barring",
    "Customer : Enquiry",
    "Customer : LAN Failure",
    "Customer : No Outage Observed",
    "Customer : Non Payment",
    "Customer : Reconnection",
    "Customer : Router Configuration",
    "Customer : Unknown",
    "Network : Bandwidth Maxing",
    "Physical Connection of Router Port",
    "Customer : Physical Router Connection",
    "Router Configuration : No downtime observed",
    "Customer : Radio / Router reboot",
    "Customer : Password Reset",
]

without_customer_issues = [
    "24 Online : Monthly Renewal",
    "24Online : Monthly Renewal",
    "Equipment : Hardware Failure : IDU",
    "Equipment : Hardware Failure : Radio",
    "Equipment : Hardware Failure : Router",
    "Equipment : Hardware Failure",
    "Equipment : Soft Failure : Radio",
    "Equipment : Soft Failure : Router",
    "Fiber Outage : Backhaul Issue",
    "Fiber Outage : FTTH : Low power signal",
    "Fiber Outage : FTTH : Loss of Signal",
    "Hardware : Cable issue",
    "Hardware : Cable Issue",
    "LOS - Radio Alignment",
    "LOS- Radio Alignment",
    "MW Transmission outages : Backhaul Issue",
    "MW Transmission Outage : Backhaul Issue",
    "MW Transmission outages : Last Mile",
    "MW Transmission Outage : Last Mile",
    "Network : Site Passive Failure",
    "Planned Activity",
    "Projects : Radio replacement",
    "Projects : Relocation",
    "Projects : Service integration",
    "Project : Service Integration",
    "Projects : Physical Survey",
    "Radio Configuration : Subscriber Module - SM",
    "Radio Link Test : Poor uploads/ downloads",
    "Radio Link Test : Radio packet drops",
    "Radio Link Test : Radio Packet Drops",
    "Radio Link Test : Poor uploads/Downloads",
    "Radio Stability : Access Point Down",
    "Radio Stability : Access Point - AP",
    "Router Configuration : Default in configuration",
    "SIP Line : ODU/5G Router",
    "SIP Line",
    "Fiber Outage : Last Mile",
    "Equipment : Software Failure",
    "Radio Configuration : Access Point - AP",
    "24Online : Outage",
    "24 Online : Outage",
    "Non Outage Related Ticket : NSA - DCN Visibility",
    "NON Outage Related Ticket : NSA - DCN Visibility : No Downtime Observed",
    "NON Outage Related Ticket : NSA - DCN Visibility : No downtime observed",
]

l1_resolvers = [
    "Customer : Device powered off",
    "Customer : Device Powered off",
    "Customer : Enquiry",
    "Customer : LAN Failure",
    "Customer : No Outage Observed",
    "Customer : Router Configuration",
    "Customer : Unknown",
    "Equipment : Soft Failure : Radio",
    "Equipment : Soft Failure : Router",
    "Network : Bandwidth Maxing",
    "Physical Connection of Router Port",
    "Customer : Physical Router Connection",
    "Radio Link Test : Poor uploads/ downloads",
    "Radio Link Test : Poor uploads/Downloads",
    "Router Configuration : No downtime observed",
    "Customer : Radio / Router reboot",
    "Customer : Password Reset",
    "Radio Configuration : Access Point - AP",
    "Non Outage Related Ticket : NSA - DCN Visibility",
    "NON Outage Related Ticket : NSA - DCN Visibility : No Downtime Observed",
    "NON Outage Related Ticket : NSA - DCN Visibility : No downtime observed",
]

l2_resolvers = [
    "24 Online : Monthly Renewal",
    "24Online : Monthly Renewal",
    "Customer : Barring",
    "Customer : Non Payment",
    "Customer : Reconnection",
    "Equipment : Hardware Failure : IDU",
    "Equipment : Hardware Failure : Radio",
    "Equipment : Hardware Failure : Router",
    "Equipment : Hardware Failure",
    "Fiber Outage : Backhaul Issue",
    "Hardware : Cable issue",
    "Hardware : Cable Issue",
    "LOS - Radio Alignment",
    "LOS- Radio Alignment",
    "MW Transmission outages : Backhaul Issue",
    "MW Transmission Outage : Backhaul Issue",
    "MW Transmission outages : Last Mile",
    "MW Transmission Outage : Last Mile",
    "Network : Site Passive Failure",
    "Planned Activity",
    "Projects : Radio replacement",
    "Projects : Relocation",
    "Projects : Service integration",
    "Project : Service Integration",
    "Projects : Physical Survey",
    "Radio Configuration : Subscriber Module - SM",
    "Radio Link Test : Radio packet drops",
    "Radio Link Test : Radio Packet Drops",
    "Radio Stability : Access Point Down",
    "Radio Stability : Access Point - AP",
    "Router Configuration : Default in configuration",
    "SIP Line : ODU/5G Router",
    "SIP Line",
    "Fiber Outage : Last Mile",
    "Fiber Outage : FTTH : Low power signal",
    "Fiber Outage : FTTH : Loss of Signal",
    "Equipment : Software Failure",
    "24Online : Outage",
    "24 Online : Outage",
]

def etl(path, country, month_date):
    # Extract
    df = pd.read_excel(path, sheet_name="Sheet1")
    # Filter Cancelled
    df_clean = df[df["Status"] != "Cancelled"].copy()

    # Convert date columns
    df_clean["Date"] = pd.to_datetime(df_clean["Date"])
    df_clean["Closed Date"] = pd.to_datetime(df_clean["Close Date"])
    df_clean["Problem End Time"] = pd.to_datetime(df_clean["Problem End Time"])
    df_clean["Problem End Time"] = df_clean["Problem End Time"].fillna(df_clean["Closed Date"])

    # Resolver + Issue Type
    df_clean["Resolver"] = np.select(
        [
            df_clean["Request Type"].str[3:].isin(l1_resolvers),
            df_clean["Request Type"].str[3:].isin(l2_resolvers),
        ],
        ["L1", "L2"],
        default=None,
    )
    df_clean["Issue Type"] = np.select(
        [
            df_clean["Request Type"].str[3:].isin(with_customer_issues),
            df_clean["Request Type"].str[3:].isin(without_customer_issues),
        ],
        ["With Customer Issues", "Without Customer Issues"],
        default=None,
    )

    # MTTR (hours)
    df_clean["MTTR"] = (df_clean["Problem End Time"] - df_clean["Date"]).dt.total_seconds() / 3600.0

    # Country subset
    df_ct = df_clean[df_clean["Request Type"].str.startswith(country)]

    # Country KPIs
    ct_total_tickets = df_ct["No."].count()
    ct_Pro_tickets = df_ct[df_ct["Request Type.1"] == "Proactive"]["No."].count()
    ct_Re_tickets = df_ct[df_ct["Request Type.1"] == "Reactive"]["No."].count()

    ct_percentage_pro = round((ct_Pro_tickets / ct_total_tickets) * 100, 2) if ct_total_tickets else 0.0
    ct_percentage_re = round((ct_Re_tickets / ct_total_tickets) * 100, 2) if ct_total_tickets else 0.0

    ct_mttr_within = df_ct[df_ct["MTTR"] <= 4]["No."].count()
    ct_mttr_outside = df_ct[df_ct["MTTR"] > 4]["No."].count()
    ct_sla_perc = round((ct_mttr_within / ct_total_tickets) * 100, 2) if ct_total_tickets else 0.0
    ct_mttr_avg = round(df_ct["MTTR"].mean(), 2) if ct_total_tickets else 0.0
    ct_without_ci = df_ct[df_ct["Issue Type"] == "Without Customer Issues"]["No."].count()
    ct_mean_without_ci = round(
        df_ct[df_ct["Issue Type"] == "Without Customer Issues"]["MTTR"].mean(), 2
    ) if ct_without_ci else 0.0

    # Filter out some specific issues
    filter_ct = df_ct[~df_ct["Request Type"].str[3:].isin(
        [
            "Customer : Device Powered off",
            "Customer : Device powered off",
            "Customer : Radio / Router reboot",
            "Network : Interference",
        ]
    )]

    fil_total_tickets = filter_ct["No."].count()
    fil_pro_tickets = filter_ct[filter_ct["Request Type.1"] == "Proactive"]["No."].count()
    fil_re_tickets = filter_ct[filter_ct["Request Type.1"] == "Reactive"]["No."].count()
    fil_percentage_pro = round((fil_pro_tickets / fil_total_tickets) * 100, 2) if fil_total_tickets else 0.0
    fil_percentage_re = round((fil_re_tickets / fil_total_tickets) * 100, 2) if fil_total_tickets else 0.0
    fil_mttr_avg = round(filter_ct["MTTR"].mean(), 2) if fil_total_tickets else 0.0
    fil_mttr_within = filter_ct[filter_ct["MTTR"] <= 4]["No."].count()
    fil_mttr_outside = filter_ct[filter_ct["MTTR"] > 4]["No."].count()
    fil_percentage_sla = round((fil_mttr_within / fil_total_tickets) * 100, 2) if fil_total_tickets else 0.0
    fil_mean_withoutcustomerissues = round(
        filter_ct[filter_ct["Issue Type"] == "Without Customer Issues"]["MTTR"].mean(), 2
    ) if filter_ct[filter_ct["Issue Type"] == "Without Customer Issues"]["No."].count() else 0.0

    fault_rate = round((fil_total_tickets / ct_total_tickets) * 100, 2) if ct_total_tickets else 0.0

    resultSet_rows = [
        ct_total_tickets,
        ct_Pro_tickets,
        ct_percentage_pro,
        ct_percentage_re,
        fault_rate,
        ct_sla_perc,
        ct_mttr_avg,
        " ",
        ct_total_tickets,
        ct_Pro_tickets,
        ct_Re_tickets,
        ct_mttr_within,
        ct_mttr_outside,
        ct_mttr_avg,
        ct_sla_perc,
        ct_mean_without_ci,
        fil_total_tickets,
        fil_pro_tickets,
        fil_re_tickets,
        fil_mttr_within,
        fil_mttr_outside,
        fil_percentage_sla,
        fil_mttr_avg,
        fil_mean_withoutcustomerissues,
    ]
    resultSet = pd.DataFrame(resultSet_rows, columns=[month_date])
    return resultSet

def update_masterfile_with_data(masterfile, df, month):
    # Put incoming month column into masterfile
    masterfile.loc[:, month] = df[month].values
    month_cols = masterfile.columns[2:]  # Assuming first 2 columns are identifiers

    fy26 = []
    for idx in range(len(masterfile)):
        vals = pd.to_numeric(masterfile.loc[idx, month_cols], errors="coerce")
        if idx in [0, 1]:
            fy26.append(vals.sum())
        elif 2 <= idx <= 7:
            fy26.append(vals.mean())
        else:
            fy26.append(None)
    masterfile["FY26"] = fy26
    return masterfile

def run_etl(input_path, output_path, month_text):
    # Build the three country DataFrames and write back into the output file
    df_ke = etl(input_path, "KE", month_text)
    df_ug = etl(input_path, "UG", month_text)
    df_sc = etl(input_path, "SC", month_text)

    sheets = {
        "KE_2026": df_ke,
        "UG_2026": df_ug,
        "SC_2026": df_sc,
    }

    # Write into existing output file sheets (replace)
    with pd.ExcelWriter(output_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        for sheet_name, df in sheets.items():
            masterfile = pd.read_excel(output_path, sheet_name=sheet_name)
            if "FY26" in masterfile.columns:
                masterfile = masterfile.drop(columns=["FY26"])
            updated = update_masterfile_with_data(masterfile, df, month_text)
            updated.to_excel(writer, index=False, sheet_name=sheet_name)

# ---------- GUI (Red/Black theme) ----------

RED = "#C62828"     # strong red
BLACK = "#111111"   # near black
WHITE = "#FFFFFF"

def resource_path(relative):
    # Helps PyInstaller find resources if needed later
    try:
        base_path = sys._MEIPASS  # type: ignore
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative)

class App:
    def __init__(self, master):
        self.master = master
        master.title("Monthly ETL Updater")

        # Dark background
        master.configure(bg=BLACK)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except:
            pass
        style.configure("TLabel", foreground=WHITE, background=BLACK, font=("Segoe UI", 10))
        style.configure("TEntry", fieldbackground=WHITE, foreground="black", padding=4)
        style.configure("TButton", foreground=WHITE, background=RED, padding=6)
        style.map("TButton",
                  foreground=[("active", WHITE)],
                  background=[("active", "#b71c1c")])

        # Title
        title = Label(master, text="Monthly ETL Updater", fg=WHITE, bg=BLACK, font=("Segoe UI", 14, "bold"))
        title.pack(pady=(10, 6))

        # INPUT: Raw file (xlsx)
        self.input_var = StringVar()
        self._row("Raw file (xlsx):", self.input_var, self.browse_input)

        # OUTPUT: Masterfile (xlsx)
        self.output_var = StringVar()
        self._row("Masterfile (xlsx):", self.output_var, self.browse_output)

        # MONTH: e.g., Feb 26
        self.month_var = StringVar()
        self._row("Month (e.g., Feb 26):", self.month_var, None)


        # Footer
                # RUN button
        btn = ttk.Button(master, text="Run", command=self.on_run)
        btn.pack(pady=(12, 10), fill=X, padx=300)

        footer = Label(master, text="Tip: Make sure both files exist and are closed in Excel.", fg="#bbbbbb", bg=BLACK)
        footer.pack(pady=(2, 10))

    def _row(self, label_text, var, browse_cmd):
        frame = Frame(self.master, bg=BLACK)
        frame.pack(fill=X, padx=12, pady=4)

        lbl = ttk.Label(frame, text=label_text)
        lbl.pack(side=LEFT)

        ent = ttk.Entry(frame, textvariable=var, width=60)
        ent.pack(side=LEFT, padx=8)

        if browse_cmd:
            btn = ttk.Button(frame, text="Browse", command=browse_cmd)
            btn.pack(side=RIGHT)

    def browse_input(self):
        path = filedialog.askopenfilename(
            title="Select raw Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if path:
            self.input_var.set(path)

    def browse_output(self):
        path = filedialog.askopenfilename(
            title="Select Masterfile.xlsx",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if path:
            self.output_var.set(path)

    def on_run(self):
        input_path = self.input_var.get().strip()
        output_path = self.output_var.get().strip()
        month_text = self.month_var.get().strip()

        # --- Simple validations like a checklist ---
        if not input_path:
            messagebox.showerror("Missing", "Please select the raw input Excel file.")
            return
        if not os.path.exists(input_path):
            messagebox.showerror("Not found", "The raw input file path does not exist.")
            return
        if not output_path:
            messagebox.showerror("Missing", "Please select the Masterfile (output).")
            return
        if not os.path.exists(output_path):
            messagebox.showerror("Not found", "The Masterfile path does not exist.")
            return
        if not month_text:
            messagebox.showerror("Missing", "Please enter month, e.g., Feb 26")
            return
        if len(month_text.split()) != 2 or len(month_text.split()[0]) != 3:
            messagebox.showerror("Format", "Month format should be like 'Feb 26'.")
            return

        try:
            run_etl(input_path, output_path, month_text)
            messagebox.showinfo("Success", f"Done! Month '{month_text}' updated.")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error", f"Something went wrong:\n{e}")

def main():
    root = Tk()
    app = App(root)
    # Fixed small window
    root.geometry("980x660")
    root.resizable(False, False)
    root.mainloop()

if __name__ == "__main__":
    main()