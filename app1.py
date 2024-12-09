import streamlit as st
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch
from sklearn.neighbors import KNeighborsClassifier
# Function to speak
def speak(str1):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str1)

# Function to take attendance
def take_attendance():
    st.write("Opening camera for attendance...")
    
    video = cv2.VideoCapture(0)
    facedetect = cv2.CascadeClassifier('face_detec/data/haarcascade_frontalface_default.xml')

    with open('face_detec/data/names.pkl', 'rb') as f:
        LABELS = pickle.load(f)
    with open('face_detec/data/faces_data.pkl', 'rb') as f:
        FACES = pickle.load(f)

    knn = KNeighborsClassifier(n_neighbors=5)
    knn.fit(FACES, LABELS)

    COL_NAMES = ['NAME', 'TIME']
    attendance_list = []

    while video.isOpened():
        ret, frame = video.read()
        if not ret:
            st.error("Failed to open the camera.")
            break

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = facedetect.detectMultiScale(gray, 1.3, 5)

        for (x, y, w, h) in faces:
            crop_img = frame[y:y + h, x:x + w, :]
            resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
            output = knn.predict(resized_img)

            ts = time.time()
            date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
            timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")
            exists = os.path.isfile("face_detec/Attendance/Attendance_" + date + ".csv")

            attendance = [str(output[0]), timestamp]
            attendance_list.append(attendance)

            if exists:
                with open("face_detec/Attendance/Attendance_" + date + ".csv", "a") as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(attendance)
            else:
                with open("face_detec/Attendance/Attendance_" + date + ".csv", "w") as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(COL_NAMES)
                    writer.writerow(attendance)

            st.write(f"Attendance marked for {output[0]} at {timestamp}")
            break  # For demonstration, stop after detecting one face
        
        # Break the loop if 'q' is pressed
        k = cv2.waitKey(1)
        if k == ord('q'):
            break
    
    video.release()
    cv2.destroyAllWindows()
    return attendance_list

# Streamlit App
st.title("Face Recognition Attendance System")

if st.button("Take Attendance"):
    attendance = take_attendance()
    if attendance:
        st.success("Attendance successfully taken!")
        st.write("Attendance Data:")
        st.table(attendance)
