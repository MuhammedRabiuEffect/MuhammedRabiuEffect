import cv2
import numpy as np
import time
import winsound
from pyzbar.pyzbar import decode
from datetime import datetime
from openpyxl import Workbook

wb = Workbook() # using the Workbook class function and methods (objects)
sheet = wb.active # using the active sheet objects

cap = cv2.VideoCapture(0)
cap.set(3, 640) #3 - width
cap.set(4, 480) #4 - heigh


def show_capture(cam):
    def scanning_time(now):
        dt = str(now).split()
        date = dt[0]
        time_ = dt[1][:-7]
        time_date = [date, time_]
        time_now = '_'.join(time_date)
        return time_now

    save_here = []

    color = np.array([35, 221, 159])
    while True:
        success, frame = cap.read()
        hsv_color = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV)
        inRange = cv2.inRange(hsv_color, color, color)

        for code in decode(frame):
            cap_qr = str(code.data.decode('utf-8')[-9:])
            if cap_qr not in save_here:
                scan_time = scanning_time(datetime.now())
                print("CODE SCANNED")
                print(f'{scan_time}............{cap_qr}')

                save_here.append(cap_qr)

                winsound.Beep(5000, 200)
                sheet.append((scan_time, cap_qr))
                print("")
                time.sleep(2)
            else:
                winsound.Beep(500, 200)
                print("CODE IS ALREADY SCANNED")
                time.sleep(2)
                print("")

        cv2.imshow('pic', frame)
        if cv2.waitKey(10) == ord('q'):
            break
    wb.save('now.xlsx')


if __name__ == '__main__':
    show_capture(cap)
