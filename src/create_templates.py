import pandas as pd

# Create data structures for various required inputs
course_data = pd.DataFrame({
    "Course Code": [],
    "Course Name": [],
    "Semester": [],
    "Department": [],
    "LTPSC": [],
    "Credits": [],
    "Instructor": [],
    "Registered Students": [],
    "Elective (Yes/No)": [],
    "Half Semester (Yes/No)": []
})

faculty_availability = pd.DataFrame({
    "Faculty Name": [],
    "Available Days": [],
    "Unavailable Time Slots": []  # e.g. "Monday 9-11, Wednesday 2-4"
})

classroom_data = pd.DataFrame({
    "Room Number": [],
    "Type": [],  # Classroom/Lab
    "Capacity": [],
    "Facilities": []  # e.g. Projector, Computers
})

student_data = pd.DataFrame({
    "Student Roll Number": [],
    "Name": [],
    "Department": [],
    "Semester": [],
    "Enrolled Courses": [],  # course codes separated by semicolons
    "Group": [],
    "Special Accommodation": []  # e.g. Disabled, None
})

exam_data = pd.DataFrame({
    "Course Code": [],
    "Exam Type": [],  # Theory/Lab/Viva
    "Exam Duration (minutes)": [],
    "Preferred Exam Date": [],
    "Alternate Exam Date": []
})

invigilator_data = pd.DataFrame({
    "Invigilator Name": [],
    "Available Days": [],
    "Unavailable Time Slots": []
})

# Write all to CSV files as templates
course_data.to_csv("course_data_template.csv", index=False)
faculty_availability.to_csv("faculty_availability_template.csv", index=False)
classroom_data.to_csv("classroom_data_template.csv", index=False)
student_data.to_csv("student_data_template.csv", index=False)
exam_data.to_csv("exam_data_template.csv", index=False)
invigilator_data.to_csv("invigilator_data_template.csv", index=False)

"CSV templates created for course, faculty, classroom, student, exam, and invigilator data."