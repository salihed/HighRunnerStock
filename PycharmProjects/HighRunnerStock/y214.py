import pandas as pd
import os
import datetime
import tkinter as tk
from tkinter import simpledialog, messagebox


def determine_shift(time_string):
    if isinstance(time_string, datetime.time):
        time_string = time_string.strftime('%H:%M:%S')
    if "08:00:00" <= time_string <= "15:59:59":
        return "8-16"
    elif "16:00:00" <= time_string <= "23:59:59":
        return "16-24"
    else:
        return "24-08"


def calculate_distance(address1, address2):
    x1, y1 = int(address1[3:6]), int(address1[6:])
    x2, y2 = int(address2[3:6]), int(address2[6:])
    distance_x = abs(x1 - x2) * 0.9
    distance_y = abs(y1 - y2) * 0.95
    return distance_x + distance_y


def calculate_and_display(date_entry):
    selected_date = date_entry.get()
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    df_y214 = pd.read_excel(desktop_path + "\\Y214.xlsx")
    df_y214["Shift"] = df_y214["Onay saati"].apply(determine_shift)

    for shift in ["8-16", "16-24", "24-08"]:
        df_filtered = df_y214[(df_y214["Onay tarihi"] == selected_date) & (df_y214["Shift"] == shift)]

        total_distance_N010 = 0
        total_distance_N020 = 0
        total_picks_N010 = 0
        total_picks_N020 = 0

        for index, row in df_filtered.iterrows():
            current_address = row["Çkş.yr.dp.adr."]
            if row["Çıkış yeri depo tipi"] == "N010":
                total_distance_N010 += (calculate_distance("01A00001", current_address) * 2)
                total_picks_N010 += 1
            elif row["Çıkış yeri depo tipi"] == "N020":
                total_distance_N020 += calculate_distance("01A00001", current_address)
                total_picks_N020 += 1

        total_distance_N010 = total_distance_N010 / 1000  # Convert to km
        total_distance_N020 = total_distance_N020 / 1000  # Convert to km

        display_result = f"{shift} Vardiyası için:\n" \
                         f"N010: Mesafe = {total_distance_N010:.2f} km, Toplam Pick = {total_picks_N010}\n" \
                         f"N020: Mesafe = {total_distance_N020:.2f} km, Toplam Pick = {total_picks_N020}\n"

        tk.messagebox.showinfo(title=f"{selected_date} - {shift}", message=display_result)


root = tk.Tk()
root.title("Mesafe Hesaplama")

tk.Label(root, text="Tarih (GG.AA.YYYY):").grid(row=0, column=0)
date_entry = tk.Entry(root)
date_entry.grid(row=0, column=1)
tk.Button(root, text="Hesapla", command=lambda: calculate_and_display(date_entry)).grid(row=1, columnspan=2)

root.mainloop()
