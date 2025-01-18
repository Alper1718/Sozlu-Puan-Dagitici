import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook
import random

FILE_NAME = "PERFORMANS ÖLÇEK sınıf listeli olan SON.xlsx"

workbook = load_workbook(FILE_NAME)
sheet = workbook.active

CRITERIA_MAX_VALUES = [40, 10, 20, 20, 10]

def distribute_points(final_points):
    points = [0] * len(CRITERIA_MAX_VALUES)

    if final_points % 5 != 0:
        raise ValueError("Final points must be a multiple of 5.")

    remaining_points = final_points
    while remaining_points > 0:
        idx = random.randint(0, len(CRITERIA_MAX_VALUES) - 1)
        if points[idx] + 5 <= CRITERIA_MAX_VALUES[idx] and remaining_points >= 5:
            points[idx] += 5
            remaining_points -= 5

    return points

def distribute_exact_points():
    return CRITERIA_MAX_VALUES.copy()

def save_to_excel(name, final_points, points):
    row_idx = names.index(name) + 8
    for col, pts in enumerate(points, start=4):
        sheet.cell(row=row_idx, column=col, value=pts)
    workbook.save(FILE_NAME)

def submit():
    try:
        for name, entry in zip(names, note_entries):
            note_text = entry.get()
            if not note_text:
                continue

            final_points = int(note_text)
            if final_points < 0 or final_points > 100:
                raise ValueError(f"Final points for {name} must be between 0 and 100.")

            if final_points == 100:
                points = distribute_exact_points()
            elif final_points % 5 == 0:
                points = distribute_points(final_points)
            else:
                raise ValueError(f"Final points for {name} must be a multiple of 5.")

            save_to_excel(name, final_points, points)

        messagebox.showinfo("Success", "All data saved successfully!")
        root.destroy()

    except ValueError as e:
        messagebox.showerror("Invalid Input", str(e))

names = []
for row in range(8, sheet.max_row + 1):
    name = sheet.cell(row=row, column=3).value
    school_number = sheet.cell(row=row, column=2).value
    if name:
        names.append(f"{school_number}  {row-7}-) {name}")

root = tk.Tk()
root.title("Performance Notes")

canvas = tk.Canvas(root)
scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
scrollable_frame = ttk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

note_entries = []

for i, name in enumerate(names):
    name_label = ttk.Label(scrollable_frame, text=name)
    name_label.grid(row=i, column=0, sticky=tk.W, pady=2)

    note_entry = ttk.Entry(scrollable_frame, width=10)
    note_entry.grid(row=i, column=1, pady=2)
    note_entries.append(note_entry)

submit_button = ttk.Button(scrollable_frame, text="Submit", command=submit)
submit_button.grid(row=len(names) + 1, column=0, columnspan=2, pady=10)

root.mainloop()
