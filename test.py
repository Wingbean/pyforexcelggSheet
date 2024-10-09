import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook
import win32com.client as win32

# ฟังก์ชันสำหรับอัปเดตข้อมูลลงในไฟล์ Excel
def update_excel(data):
    try:
        # โหลดไฟล์ Excel
        file_path = 'C:\Users\makhu\OneDrive\Py\pyforexcelggSheet/ใบรับรองแพทย์ 5 โรค.xlsx'
        workbook = load_workbook(filename=file_path)
        sheet = workbook.active
        
        # ใส่ข้อมูลในเซลล์ที่ต้องการ
        sheet['Z3'] = data[0]
        sheet['Z4'] = data[1]
        sheet['Y5'] = data[2]
        sheet['Z5'] = data[3]
        sheet['AA5'] = data[4]
        sheet['AB5'] = data[5]
        sheet['AC5'] = data[6]
        sheet['Z6'] = data[7]
        sheet['Z7'] = data[8]
        sheet['Z8'] = data[9]
        sheet['Z9'] = data[10]
        sheet['Z10'] = data[11]
        sheet['Z11'] = data[12]

        # บันทึกไฟล์
        workbook.save(file_path)
        workbook.close()
        messagebox.showinfo("Success", "Data saved successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save data: {e}")

# ฟังก์ชันสำหรับสั่งพิมพ์หน้าแรกของไฟล์ Excel
def print_excel():
    try:
        file_path = 'C:\Users\makhu\OneDrive\Py\pyforexcelggSheet/ใบรับรองแพทย์ 5 โรค.xlsx'
        excel = win32.Dispatch('Excel.Application')
        wb = excel.Workbooks.Open(file_path)
        ws = wb.Sheets(1)
        
        # สั่งพิมพ์หน้าแรกของไฟล์ Excel
        ws.PrintOut(From=1, To=1)
        wb.Close(SaveChanges=False)
        excel.Quit()
        
        messagebox.showinfo("Success", "Printed successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to print: {e}")

# ฟังก์ชันเมื่อกดปุ่ม Submit
def submit():
    data = [
        entry1.get(), entry2.get(), entry3.get(), entry4.get(),
        entry5.get(), entry6.get(), entry7.get(), entry8.get(),
        entry9.get(), entry10.get(), entry11.get(), entry12.get(),
        entry13.get()
    ]
    update_excel(data)

# สร้างหน้าต่าง GUI
root = tk.Tk()
root.title("Excel Data Entry")
root.geometry("400x600")

# สร้างช่องกรอกข้อมูล
labels = []
entries = []
for i in range(1, 14):
    label = tk.Label(root, text=f"Input {i}:")
    label.pack()
    labels.append(label)
    
    entry = tk.Entry(root)
    entry.pack()
    entries.append(entry)

(entry1, entry2, entry3, entry4, entry5, entry6, entry7, entry8, 
 entry9, entry10, entry11, entry12, entry13) = entries

# สร้างปุ่ม Submit และ Print
submit_button = tk.Button(root, text="Submit", command=submit)
submit_button.pack(pady=10)

print_button = tk.Button(root, text="Print", command=print_excel)
print_button.pack(pady=10)

# เริ่ม GUI loop
root.mainloop()
