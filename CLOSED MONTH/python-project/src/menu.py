# menu.py

import os
import tkinter as tk
from tkinter import messagebox, filedialog
from custom_reports import create_custom_report
from acceptance_transfer_acts import AcceptanceTransferActCreator
from openpyxl import load_workbook
import pandas as pd

def open_excel_file():
    file_path = filedialog.askopenfilename(
        title="Оберіть Excel-файл",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        try:
            os.startfile(file_path)
            messagebox.showinfo("Успіх", "Файл відкрито.")
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося відкрити файл: {e}")

def create_monthly_agreement():
    messagebox.showinfo("Місячна угода", "Тут буде логіка створення місячної угоди.")

def create_acceptance_transfer_act():
    file_path = filedialog.askopenfilename(
        title="Оберіть Excel-файл з вихідними даними",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not file_path:
        messagebox.showinfo("Відміна", "Файл не обрано.")
        return

    try:
        wb = load_workbook(file_path, data_only=True)
        xls = pd.ExcelFile(file_path)
        delivery_df = xls.parse('Delivery&JAO schedule', header=None)
        template_path = os.path.join(
            os.path.dirname(os.path.dirname(__file__)), 'templates', 'Acceptance Certificate_FORM.docx'
        )

        act_creator = AcceptanceTransferActCreator()
        act_creator.create_act(wb, delivery_df, is_import=True, template_path=template_path)
        act_creator.create_act(wb, delivery_df, is_import=False, template_path=template_path)
        messagebox.showinfo("Успіх", "Акти приймання-передачі сформовано.")
    except Exception as e:
        messagebox.showerror("Помилка", f"Не вдалося створити акт: {e}")

def main_menu():
    root = tk.Tk()
    root.title("Головне меню")

    tk.Label(root, text="Оберіть дію:", font=("Arial", 14)).pack(pady=10)

    tk.Button(root, text="Створити розрахунок для митниці", width=40, command=create_custom_report).pack(pady=5)
    tk.Button(root, text="Створити місячну угоду про поставку", width=40, command=create_monthly_agreement).pack(pady=5)
    tk.Button(root, text="Створити акт приймання-передачі", width=40, command=create_acceptance_transfer_act).pack(pady=5)
    tk.Button(root, text="Відкрити Excel-файл з вихідними даними", width=40, command=open_excel_file).pack(pady=5)
    tk.Button(root, text="Вийти", width=40, command=root.destroy).pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main_menu()