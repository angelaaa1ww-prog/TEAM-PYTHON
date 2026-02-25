import tkinter as tk
from tkinter import messagebox
from student import Student  # <- make sure student.py is in the same folder

students = []

def register_student():
    fn = entry_first.get()
    ln = entry_last.get()
    phone = entry_phone.get()
    sid = entry_id.get()
    course = entry_course.get()

    if not fn or not ln or not phone or not sid or not course:
        messagebox.showerror("Error", "All fields are required")
        return

    student = Student(fn, ln, phone, sid, course)
    students.append(student)
    listbox.insert(tk.END, student.get_details())

    entry_first.delete(0, tk.END)
    entry_last.delete(0, tk.END)
    entry_phone.delete(0, tk.END)
    entry_id.delete(0, tk.END)
    entry_course.delete(0, tk.END)

root = tk.Tk()
root.title("School Management")
root.geometry("600x500")

tk.Label(root, text="First Name").pack()
entry_first = tk.Entry(root)
entry_first.pack()

tk.Label(root, text="Last Name").pack()
entry_last = tk.Entry(root)
entry_last.pack()

tk.Label(root, text="Phone").pack()
entry_phone = tk.Entry(root)
entry_phone.pack()

tk.Label(root, text="Student ID").pack()
entry_id = tk.Entry(root)
entry_id.pack()

tk.Label(root, text="Course").pack()
entry_course = tk.Entry(root)
entry_course.pack()

tk.Button(root, text="Register Student", command=register_student).pack(pady=10)

listbox = tk.Listbox(root, width=80)
listbox.pack(pady=10)

root.mainloop()