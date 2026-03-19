from employee_database_class_excel import EmployeeDatabase

db = EmployeeDatabase()

db.insert_employee("John", "Doe", "1980-01-15", "2020-06-10", "Software Engineer")

employee = db.get_employee("John", "Doe")
if employee:
    print(employee);
else:
    print("Employee not found.")
#adapter is able to run smoothly and
# not hit error of not finding the employee
# works with empty and non-empty sheet
# will create new headers for empty sheet
# wont create extra for non-empty sheet

db.insert_employee("Jane", "Doe", "1985-10-15", "2019-10-09", "Scientist")
db.insert_employee("Frank", "Matthew", "1955-10-15", "2000-01-01", "Janitor")
db.insert_employee("Rick", "Fox", "2001-10-15", "2025-02-14", "Intern")

employees = db.getAllEmp()
print(employees)
# works will return all emplyees in a list with id, name,
# birthdate, hire date, and job title

print(db.readSheet())
# read sheet works properly
# will output a table looking thing in the console

db.close()
# i can only assume this works
