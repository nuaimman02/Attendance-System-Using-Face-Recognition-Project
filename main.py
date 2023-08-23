import cv2
import numpy as np
import face_recognition
import os
import time
import csv
from win32com.client import Dispatch
from datetime import datetime, date

datetoday = date.today().strftime("%m_%d_%y")
datetoday2 = date.today().strftime("%d-%B-%Y")

attendance_path = "C:\\Users\\User\\Documents\\Programming Codes\\Python\\PyCharmProject\\pythonProject\\CSC3600_AttendanceSystemProject\\attendance.csv"

path = 'Training Images'
images = []
classNames = []
myList = os.listdir(path)
print(myList)
for cl in myList:
    curImg = cv2.imread(f'{path}/{cl}')
    images.append(curImg)
    classNames.append(os.path.splitext(cl)[0])
print(classNames)

def speak(str1):
    speak = Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)

def findEncodings(images):
    encodeList = []


    for img in images:
        img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        encode = face_recognition.face_encodings(img)[0]
        encodeList.append(encode)
    return encodeList

encodeListKnown = findEncodings(images)
print('Encoding Complete')

cap = cv2.VideoCapture(0)
img = cap.read()

while True:
    success, img = cap.read()
    # img = captureScreen()
    imgS = cv2.resize(img, (0, 0), None, 0.25, 0.25)
    imgS = cv2.cvtColor(imgS, cv2.COLOR_BGR2RGB)

    facesCurFrame = face_recognition.face_locations(imgS)
    encodesCurFrame = face_recognition.face_encodings(imgS, facesCurFrame)

    for encodeFace, faceLoc in zip(encodesCurFrame, facesCurFrame):
        matches = face_recognition.compare_faces(encodeListKnown, encodeFace)
        faceDis = face_recognition.face_distance(encodeListKnown, encodeFace)
        # print(faceDis)
        matchIndex = np.argmin(faceDis)

        now = datetime.now()
        dtString = now.strftime("%d/%m/%Y,%H:%M:%S")

        exist=os.path.isfile(attendance_path)

        if matches[matchIndex]:
            name = classNames[matchIndex].upper()
            # print(name)
            y1, x2, y2, x1 = faceLoc
            y1, x2, y2, x1 = y1 * 4, x2 * 4, y2 * 4, x1 * 4
            cv2.rectangle(img, (x1, y1), (x2, y2), (0, 255, 0), 2)
            cv2.rectangle(img, (x1, y2 - 35), (x2, y2), (0, 255, 0), cv2.FILLED)
            cv2.putText(img, name, (x1 + 6, y2 - 6), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 2)

    cv2.putText(img, "Press Enter when face detected", (30, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 0, 20), 2, cv2.LINE_AA)
    imgS = img
    cv2.imshow("Attendance Camera", img)
    k = cv2.waitKey(1)

    if k == 13:
        speak("Attendance Taken..")
        time.sleep(5)

        if exist:
            with open(attendance_path, 'r+') as f:
                myDataList = f.readlines()
                nameList = []

                for line in myDataList:
                    entry = line.split(',')
                    nameList.append(entry[0])
                    f.writelines(f'\n{name},{dtString}')
        else:
            with open(attendance_path, 'a+') as f:
                myDataList = f.readlines()
                nameList = []
                f.writelines(['NAME, DATE, TIME'])

                for line in myDataList:
                    entry = line.split(',')
                    nameList.append(entry[0])
                    f.writelines(f'\n{name},{dtString}')

    if k == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()