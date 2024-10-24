Face Recognition Using Python and Excel Integration
This project demonstrates a face recognition system developed using Python, OpenCV for computer vision, and the face_recognition package for recognizing faces. The system utilizes an Excel file (managed using the openpyxl library) as a simple database to store and update user details and face recognition results.

Features
Face Recognition: Uses the face_recognition package to detect and recognize faces from live camera feed or stored images.
OpenCV Integration: Handles the computer vision tasks, such as capturing images and displaying results.
Excel File as Database: Two Excel files are used:
Users.xlsx: Stores user details (e.g., name, ID, and face encoding).
Attendance.xlsx: Tracks attendance based on recognized faces.
Automated Updates: The system automatically updates the Excel files when a face is recognized or new data is registered.
Project Structure
face_recognition.py: The main Python script to run the face recognition system.
Users.xlsx: Stores user information such as names, unique IDs, and face encodings.
Attendance.xlsx: Logs attendance, marking recognized users along with the date and time.
Requirements
Python 3.x
Packages:
face_recognition
opencv-python
openpyxl
numpy (for numerical operations)
pygame (for present sound)
datetime 

Future Improvements
Switch from Excel to a more robust database like SQLite or MySQL for better scalability.
Implement a GUI using Tkinter to handle user registration and view attendance logs more easily.
Note: use numpy 1.26.4
