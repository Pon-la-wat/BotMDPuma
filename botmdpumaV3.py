import re
import cv2
# import time
import pyodbc
import numpy as np
import pytesseract
from PIL import ImageGrab
from playsound import playsound
import pygame
import keyboard

# ตั้งค่า pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# การตั้งค่าการเชื่อมต่อฐานข้อมูล SQL Server
conn = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.50.253;DATABASE=DtradeProduction;UID=sa;PWD=Sql4116!')
cursor = conn.cursor()

def play_sound(filepath):
    pygame.mixer.init()
    pygame.mixer.music.load(filepath)
    pygame.mixer.music.play()
    while pygame.mixer.music.get_busy() == True:
        continue

def capture_screen(bbox=(0, 200, 2000, 950)):
    """จับภาพหน้าจอ"""
    screen = np.array(ImageGrab.grab(bbox=bbox))
    screen_gray = cv2.cvtColor(screen, cv2.COLOR_BGR2GRAY)
    # cv2.imshow('Screen', screen_gray)
    return screen_gray

def read_barcode():
    """อ่านบาร์โค้ดจากสแกนเนอร์ที่ทำงานเหมือนแป้นพิมพ์"""
    scanner = ''
    while True:
        event = keyboard.read_event()
        if event.event_type == keyboard.KEY_DOWN and event.name != 'enter':
            scanner += event.name
        elif event.event_type == keyboard.KEY_DOWN and event.name == 'enter':
            break
    return scanner

def realtime_ocr():
    while True:
        scanner = read_barcode()
        if scanner:
            current_screen = capture_screen()
            text = pytesseract.image_to_string(current_screen, lang='eng')
            qty = re.search(r"Quantity\s+(\d+)", text)
            packageQty = re.search(r"Package Qty\s+(\d+)", text)
            quantity = int(qty.group(1)) if qty else int(packageQty.group(1)) if packageQty else None
            
            scanned = int(re.search(r"Scanned\s+(\d+)", text).group(1)) if re.search(r"Scanned\s+(\d+)", text) else None
            barcode = str(re.search(r"LPN:\s*(\d+)", text).group(1)) if re.search(r"LPN:\s*(\d+)", text) else None

            # print(f'quantity: {quantity}')
            # print(f'scanned: {scanned}')
            # print(f'barcode: {barcode}')
            # print(f'Barcode scanned: {barcode}')
            # print('Processing OCR...')

            # แสดงผลข้อความด้วยการพิมพ์แบบปลอดภัยเพื่อหลีกเลี่ยงข้อผิดพลาดการเข้ารหัส
            try:
                if quantity == scanned and (quantity != None and scanned != None and barcode != None):
                    print('quantity: '+str(quantity)+' == scanned: '+str(scanned)+' = barcode: '+str(barcode))
                    # สร้างและเรียกใช้คำสั่ง SQL สำหรับการค้นหา
                    query = '''SELECT pb.PKInsNo, pb.JobNo, pb.JobDetId, pb.PPTxnId, pb.PrePackTxnId, pb.TxnId, pb.Type, pb.NewBarcode, pb.CartonQty, pb.CartonNo, pb.Status, ip.ColorId, ip.SizeId, ip.InseamId, ip.Qty
                    FROM Tx_SP_JobDetInstPackBarcode AS pb INNER JOIN Tx_SP_JobDetInstPack AS ip ON pb.PKInsNo = ip.PKInsNo AND pb.JobNo = ip.JobNo AND pb.JobDetId = ip.JobDetId AND pb.TxnId = ip.TxnId
                    WHERE (pb.Status IS NULL) AND (pb.NewBarcode = ?)'''
                    cursor.execute(query, (barcode,))
                    
                    # ตรวจสอบผลลัพธ์
                    row = cursor.fetchone()
                    if row:
                        cursor.execute("SELECT COUNT(*) FROM Tx_SP_ScanPackDet WHERE Barcode = ?", (barcode,))
                        result = cursor.fetchone()
                        if result[0] == 0:
                            # ถ้าไม่มีข้อมูลที่ซ้ำกัน, ทำการ INSERT
                            cursor.execute('''INSERT INTO Tx_SP_ScanPackDet (Type, Barcode, TxnId, PackTypeId, JobNo, JobDetId, PLTxnId, PPTxnId, PrePackTxnId, ColorId, SizeId, InseamId, TargetQty, ActualQty, Qty,
                            SPDate, PackListNo, CartonNo, CreateUser, CreateDateTime)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, GETDATE(), ?, ?, 'BOT', GETDATE())''',
                            (row[6], row[7], row[5], 1, row[1], row[2], 1, row[3], row[4], row[11], row[12], row[13], row[14], row[14], row[14], row[0], row[9]))
                            # ทำการ UPDATE Status
                            cursor.execute("UPDATE Tx_SP_JobDetInstPackBarcode SET Status = 'PK' WHERE NewBarcode = ? AND Status IS NULL", (barcode))
                            conn.commit()
                            
                            print('update and insert success!')
                            play_sound('../sounds/alert.mp3')
                    else:
                        print("duplicate information.")
                else:
                    print('quantity: '+str(quantity)+' != scanned: '+str(scanned)+' = barcode: '+str(barcode))
            except UnicodeEncodeError:
                print(text.encode('ascii', 'replace').decode('ascii'))

        # หน่วงเวลาเพื่อลดการใช้งานทรัพยากร
        # time.sleep(1)

        # ออกจากลูปด้วยการกด 'q' หรือเงื่อนไขอื่นๆ
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    # ปิดการเชื่อมต่อ
    cursor.close()
    conn.close()
    cv2.destroyAllWindows()

# เริ่มต้นการทำงาน
realtime_ocr()