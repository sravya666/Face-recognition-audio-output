import face_recognition
import cv2
import numpy as np
import os
import pandas as pd
import time
from utils import append_df_to_excel
import pyttsx3

engine = pyttsx3.init()

import datetime


# Get a reference to webcam #0 (the default one)
video_capture = cv2.VideoCapture(0)


# Initialize some variables
known_face_encodings = []
known_face_roll_no = []
face_locations = []
face_encodings = []
face_names = []
process_this_frame = True
attendance_record = set([])
roll_record = {}

# Rows in log file
name_col = []
roll_no_col = []
time_col = []

df = pd.read_excel("student_db" + os.sep + "people_db.xlsx")

for key, row in df.iterrows():
    roll_no = row['roll_no']
    name = row['name']
    image_path = row['image']
    roll_record[roll_no] = name
    student_image = face_recognition.load_image_file(
        "student_db" + os.sep + image_path)
    student_face_encoding = face_recognition.face_encodings(student_image)[0]
    known_face_encodings.append(student_face_encoding)
    known_face_roll_no.append(roll_no)


while True:
    # Grab a single frame of video
    ret, frame = video_capture.read()

    # Resize frame of video to 1/4 size for faster face recognition processing
    small_frame = cv2.resize(frame, (0, 0), fx=1, fy=1)

    # Convert the image from BGR color (which OpenCV uses) to RGB color (which face_recognition uses)
    rgb_small_frame = small_frame[:, :, ::-1]

    # Only process every other frame of video to save time
    if process_this_frame:
        # Find all the faces and face encodings in the current frame of video
        face_locations = face_recognition.face_locations(rgb_small_frame)
        face_encodings = face_recognition.face_encodings(
            rgb_small_frame, face_locations)

        face_names = []
        for face_encoding in face_encodings:
            # See if the face is a match for the known face(s)
            matches = face_recognition.compare_faces(
                known_face_encodings, face_encoding, tolerance=0.5)
            name = "Unknown"

            # Or instead, use the known face with the smallest distance to the new face
            face_distances = face_recognition.face_distance(
                known_face_encodings, face_encoding)
            best_match_index = np.argmin(face_distances)
            if matches[best_match_index]:
                roll_no = known_face_roll_no[best_match_index]
                # add this to the log
                name = roll_record[roll_no]
                if roll_no not in attendance_record:
                    attendance_record.add(roll_no)
                    x = datetime.datetime.now()
                    print(name, roll_no, x)
                    name_col.append(name)
                    roll_no_col.append(roll_no)
                    curr_time = time.localtime()
                    curr_clock = time.strftime("%H:%M:%S", curr_time)
                    time_col.append(curr_clock)
                    

            engine.say(name)
            engine.runAndWait()


            face_names.append(name)
            x = datetime.datetime.now()
            

    process_this_frame = not process_this_frame

    # Display the results
    for (top, right, bottom, left), name in zip(face_locations, face_names):

        # Draw a box around the face
        cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2)

        # Draw a label with a name below the face
        cv2.rectangle(frame, (left, bottom - 35),
                      (right, bottom), (0, 255, 0), cv2.FILLED)
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(frame, name, (left + 6, bottom - 6),
                    font, 1.0, (255, 255, 255), 1)

    # Display the resulting image
    cv2.imshow('Video', frame)

    # Hit 'q' on the keyboard to quit!
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break



# Printing to log file

data = {'Name': name_col, 'Roll No.': roll_no_col, 'Time': time_col}


# Release handle to the webcam
video_capture.release()
cv2.destroyAllWindows()
