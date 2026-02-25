# Author : Amani Angela
# Date : 23/02/2026 
# Program to define a Person class for a school management system   




from person import Person

class Student(Person):
    def __init__(self, first_name, last_name, phone, student_id, course):
        super().__init__(first_name, last_name, phone)
        self.student_id = student_id
        self.course = course

    def get_details(self):
        return f"{self.first_name} {self.last_name} | Phone: {self.phone} | ID: {self.student_id} | Course: {self.course}"