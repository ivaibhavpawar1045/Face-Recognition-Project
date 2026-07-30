# 🎯 Face Recognition Attendance System

![Python](https://img.shields.io/badge/Python-3.x-blue)
![OpenCV](https://img.shields.io/badge/OpenCV-4.x-green)
![Streamlit](https://img.shields.io/badge/Streamlit-Framework-red)
![License](https://img.shields.io/badge/License-MIT-yellow)
![Status](https://img.shields.io/badge/Status-Active-brightgreen)

A real-time Face Recognition Attendance System built using **Python**, **OpenCV**, **face_recognition**, and **Streamlit**. The application recognizes registered users through a webcam and automatically records attendance with the current date and time in both CSV and Excel files.

---

## 🚀 Features

- ✅ Real-time face detection using OpenCV
- ✅ Face recognition using the `face_recognition` library
- ✅ Automatic attendance marking
- ✅ Stores attendance in CSV and Excel formats
- ✅ Prevents duplicate attendance for the same day
- ✅ Saves images of unknown faces
- ✅ Simple and interactive Streamlit interface

---

## 🛠️ Tech Stack

- Python
- Streamlit
- OpenCV
- face_recognition
- NumPy
- Pandas
- OpenPyXL
- CSV & Excel

---

## 📁 Project Structure

```text
Face-Recognition-Project/
│
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
├── known_faces/
├── unknown_faces/
├── attendance.csv
└── attendance.xlsx
```

---

## ⚙️ Installation

### 1. Clone the repository

```bash
git clone https://github.com/ivaibhavpawar1045/Face-Recognition-Project.git
```

### 2. Open the project folder

```bash
cd Face-Recognition-Project
```

### 3. Install all required packages

```bash
pip install -r requirements.txt
```

### 4. Run the application

```bash
streamlit run app.py
```

---

## 📌 Usage

1. Add images of registered users inside the `known_faces` folder.
2. Start the application.
3. Click **Take Attendance**.
4. The webcam will detect and recognize faces.
5. Attendance is automatically recorded in:
   - `attendance.csv`
   - `attendance.xlsx`
6. Unknown faces are saved inside the `unknown_faces` folder.

---

## 📊 Output

- Attendance with Name, Date, and Time
- CSV attendance report
- Excel attendance report
- Unknown face images

---

## 🔮 Future Improvements

- Anti-spoofing (Proxy Detection)
- Face Registration Module
- Database Integration
- Admin Dashboard
- Cloud Deployment
- Multi-user Authentication

---

## 👨‍💻 Author

**Vaibhav Pawar**

GitHub: https://github.com/ivaibhavpawar1045

LinkedIn: https://www.linkedin.com/in/vaibhav-pawar1045

---

⭐ If you found this project useful, consider giving it a Star on GitHub.
