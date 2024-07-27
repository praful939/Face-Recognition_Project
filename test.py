from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime


from win32com.client import Dispatch

def speak(str1):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)

video=cv2.VideoCapture(0)
facedetect=cv2.CascadeClassifier(r'C:\Users\prafu\OneDrive\Desktop\face_dupli\face_recognition_project\data\haarcascade_frontalface_default.xml')

with open(r'C:\Users\prafu\OneDrive\Desktop\face_dupli\face_recognition_project\data\names.pkl', 'rb') as w:
    LABELS=pickle.load(w)
with open(r'C:\Users\prafu\OneDrive\Desktop\face_dupli\face_recognition_project\data\faces_data.pkl', 'rb') as f:
    FACES=pickle.load(f)

print('Shape of Faces matrix --> ', FACES.shape)

knn=KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

imgBackground=cv2.imread(r'C:\Users\prafu\OneDrive\Desktop\face_dupli\face_recognition_project\background.png')

COL_NAMES = ['NAME', 'TIME']
DISTANCE_THRESHOLD = 3400
date = datetime.now().strftime("%d-%m-%Y")
attendance = ["unknown", "unknown"]

while True:
    ret,frame=video.read()
    gray=cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces=facedetect.detectMultiScale(gray, 1.3 ,5)
    output_label = ""
    for (x,y,w,h) in faces:
        crop_img=frame[y:y+h, x:x+w, :]
        resized_img=cv2.resize(crop_img, (50,50)).flatten().reshape(1,-1)
        distances, indices = knn.kneighbors(resized_img)
        ts=time.time()
        min_distance = np.min(distances)
        date=datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp=datetime.fromtimestamp(ts).strftime("%H:%M-%S")
        exist=os.path.isfile(r'C:\Users\prafu\OneDrive\Desktop\face_dupli\face_recognition_project\Attendance\Attendance_' + date + ".csv")
        cv2.rectangle(frame, (x,y), (x+w, y+h), (0,0,255), 1)
        cv2.rectangle(frame,(x,y),(x+w,y+h),(50,50,255),2)
        cv2.rectangle(frame,(x,y-40),(x+w,y),(50,50,255),-1)
        if min_distance > DISTANCE_THRESHOLD:
            # Person is unknown
            output_label = "unknown"
            print(min_distance)
        else:
            # Get the predicted label
            output_label = knn.predict(resized_img)[0]


        cv2.putText(frame, str(output_label), (x,y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255,255,255), 1)
        cv2.rectangle(frame, (x,y), (x+w, y+h), (50,50,255), 1)
        attendance=[str(output_label), str(timestamp)]
    imgBackground[162:162 + 480, 55:55 + 640] = frame
    cv2.imshow("Frame",imgBackground)
    k=cv2.waitKey(1)
    if output_label != "unknown":
        if k==ord('o'):
            speak("Attendance Taken..")
            time.sleep(2)
            if exist:
                with open(r'C:\Users\prafu\OneDrive\Desktop\face_dupli\face_recognition_project\Attendance\Attendance_' + date + ".csv", "+a") as csvfile:
                    writer=csv.writer(csvfile)
                    writer.writerow(attendance)
                csvfile.close()
            else:
                with open(r'C:\Users\prafu\OneDrive\Desktop\face_dupli\face_recognition_project\Attendance\Attendance_' + date + ".csv", "+a") as csvfile:
                    writer=csv.writer(csvfile)
                    writer.writerow(COL_NAMES)
                    writer.writerow(attendance)
                csvfile.close()
    #if output_label == "unknown":
     #   speak("Attendance Not Taken..")
    if k==ord('q'):
        break
video.release()
cv2.destroyAllWindows()