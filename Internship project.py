from tkinter import *
from tkinter import messagebox
import xlsxwriter

def calculate_results():
    roll_no = int(rollno_entry.get())
    name = name_entry.get()
    subject1 = int(sub1_entry.get())
    subject2 = int(sub2_entry.get())
    subject3 = int(sub3_entry.get())

    total = subject1 + subject2 + subject3
    average = round(total / 3, 2)

    total_var.set(str(total))
    average_var.set(str(average))

    workbook = xlsxwriter.Workbook('student_results.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})

    worksheet.write('A1', 'Roll_No', bold)
    worksheet.write('B1', 'Name', bold)
    worksheet.write('C1', 'Subject1', bold)
    worksheet.write('D1', 'Subject2', bold)
    worksheet.write('E1', 'Subject3', bold)
    worksheet.write('F1', 'Total', bold)
    worksheet.write('G1', 'Average', bold)

    worksheet.write('A2', roll_no)
    worksheet.write('B2', name)
    worksheet.write('C2', subject1)
    worksheet.write('D2', subject2)
    worksheet.write('E2', subject3)
    worksheet.write('F2', total)
    worksheet.write('G2', average)

    workbook.close()

    messagebox.showinfo("CONFIRMATION", "COMPLETED")

root = Tk()
root.title("Result")
root.geometry("400x350")

title_label = Label(root, text="K.S Institute of Technology",
                    font=("Helvetica", 14, "bold"), pady=10)
title_label.pack()

frame = Frame(root)
frame.pack(pady=10)

Label(frame, text="Roll No", font=("Helvetica", 10)).grid(row=0, column=0)
rollno_entry = Entry(frame, font=("Helvetica", 10))
rollno_entry.grid(row=0, column=1, padx=5, pady=5)

Label(frame, text="Name", font=("Helvetica", 10)).grid(row=1, column=0)
name_entry = Entry(frame, font=("Helvetica", 10))
name_entry.grid(row=1, column=1, padx=5, pady=5)

Label(frame, text="Subject 1", font=("Helvetica", 10)).grid(row=2, column=0)
sub1_entry = Entry(frame, font=("Helvetica", 10))
sub1_entry.grid(row=2, column=1, padx=5, pady=5)

Label(frame, text="Subject 2", font=("Helvetica", 10)).grid(row=3, column=0)
sub2_entry = Entry(frame, font=("Helvetica", 10))
sub2_entry.grid(row=3, column=1, padx=5, pady=5)

Label(frame, text="Subject 3", font=("Helvetica", 10)).grid(row=4, column=0)
sub3_entry = Entry(frame, font=("Helvetica", 10))
sub3_entry.grid(row=4, column=1, padx=5, pady=5)

total_var = StringVar()
average_var = StringVar()

Label(frame, text="Total", font=("Helvetica", 10)).grid(row=5, column=0)
Entry(frame, textvariable=total_var, font=("Helvetica", 10), state='readonly').grid(row=5, column=1)

Label(frame, text="Average", font=("Helvetica", 10)).grid(row=6, column=0)
Entry(frame, textvariable=average_var, font=("Helvetica", 10), state='readonly').grid(row=6, column=1)

Button(root, text="Calculate", command=calculate_results).pack(pady=10)

root.mainloop()