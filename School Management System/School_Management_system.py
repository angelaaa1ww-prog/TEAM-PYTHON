# Name : Angela Amani
# Date : 24/02/2026
# Program to define a Person class for a school management system

import customtkinter as ctk
from tkinter import messagebox
import openpyxl

# ================== Data Storage ==================
students = []  # Local storage

# ================== Functions ==================

def register_student():
    student_id = entry_id.get()
    name = entry_name.get()
    course = entry_course.get()
    phone = entry_phone.get()

    if not student_id or not name or not course or not phone:
        messagebox.showerror("Error", "All fields are required")
        return

    for s in students:
        if s["id"] == student_id:
            messagebox.showerror("Error", "Student ID already exists")
            return

    students.append({"id": student_id, "name": name, "course": course, "phone": phone})
    messagebox.showinfo("Success", f"Student {name} registered successfully!")

    entry_id.delete(0, "end")
    entry_name.delete(0, "end")
    entry_course.delete(0, "end")
    entry_phone.delete(0, "end")

def assign_new_course():
    def update_course():
        sid = entry_assign_id.get()
        new_course = entry_assign_course.get()
        for s in students:
            if s["id"] == sid:
                s["course"] = new_course
                messagebox.showinfo("Success", f"Course updated to {new_course}")
                assign_window.destroy()
                return
        messagebox.showerror("Error", "Student ID not found")

    assign_window = ctk.CTkToplevel(root)
    assign_window.title("Assign New Course")
    assign_window.geometry("300x180")

    ctk.CTkLabel(assign_window, text="Student ID").pack(pady=5)
    entry_assign_id = ctk.CTkEntry(assign_window, placeholder_text="Enter Student ID")
    entry_assign_id.pack(pady=5)

    ctk.CTkLabel(assign_window, text="New Course").pack(pady=5)
    entry_assign_course = ctk.CTkEntry(assign_window, placeholder_text="Enter New Course")
    entry_assign_course.pack(pady=5)

    ctk.CTkButton(assign_window, text="Update Course", command=update_course, corner_radius=8).pack(pady=10)

def export_to_excel():
    if not students:
        messagebox.showerror("Error", "No data to export")
        return

    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Students"
    sheet.append(["ID", "Name", "Course", "Phone"])

    for s in students:
        sheet.append([s["id"], s["name"], s["course"], s["phone"]])

    wb.save("students.xlsx")
    messagebox.showinfo("Success", "Data exported to students.xlsx")

# ================== GUI Setup ==================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.geometry("500x500")
root.title("School Management System")

# Top frame
top_frame = ctk.CTkFrame(root)
top_frame.pack(fill="x", padx=20, pady=10)

title_label = ctk.CTkLabel(top_frame, text="School Management System", font=("Arial", 22, "bold"))
title_label.pack(pady=10)

# Form frame
form_frame = ctk.CTkFrame(root)
form_frame.pack(fill="x", padx=20, pady=10)

ctk.CTkLabel(form_frame, text="Student ID").grid(row=0, column=0, padx=10, pady=5, sticky="e")
entry_id = ctk.CTkEntry(form_frame, placeholder_text="Enter ID")
entry_id.grid(row=0, column=1, padx=10, pady=5)

ctk.CTkLabel(form_frame, text="Name").grid(row=1, column=0, padx=10, pady=5, sticky="e")
entry_name = ctk.CTkEntry(form_frame, placeholder_text="Enter Name")
entry_name.grid(row=1, column=1, padx=10, pady=5)

ctk.CTkLabel(form_frame, text="Course").grid(row=2, column=0, padx=10, pady=5, sticky="e")
entry_course = ctk.CTkEntry(form_frame, placeholder_text="Enter Course")
entry_course.grid(row=2, column=1, padx=10, pady=5)

ctk.CTkLabel(form_frame, text="Phone").grid(row=3, column=0, padx=10, pady=5, sticky="e")
entry_phone = ctk.CTkEntry(form_frame, placeholder_text="Enter Phone")
entry_phone.grid(row=3, column=1, padx=10, pady=5)

# Buttons frame
button_frame = ctk.CTkFrame(root)
button_frame.pack(fill="x", padx=20, pady=10)

ctk.CTkButton(button_frame, text="Register Student", command=register_student, width=180, height=35, corner_radius=8).grid(row=0, column=0, padx=10, pady=5)
ctk.CTkButton(button_frame, text="Assign Course", command=assign_new_course, width=180, height=35, corner_radius=8).grid(row=0, column=1, padx=10, pady=5)
ctk.CTkButton(button_frame, text="Export to Excel", command=export_to_excel, width=180, height=35, corner_radius=8).grid(row=1, column=0, columnspan=2, pady=10)

root.mainloop()