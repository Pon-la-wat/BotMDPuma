import re
import cv2
import pyodbc
import numpy as np
import pytesseract
from PIL import ImageGrab
from playsound import playsound

# ตั้งค่า pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# การตั้งค่าการเชื่อมต่อฐานข้อมูล SQL Server
conn = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.50.253;DATABASE=DtradeProduction;UID=sa;PWD=Sql4116!')
cursor = conn.cursor()

def realtime_ocr():
    while True:
        # จับภาพหน้าจอ
        screen = np.array(ImageGrab.grab(bbox=None))  # bbox specifies specific region (bbox=None captures the entire screen)
        screen_gray = cv2.cvtColor(screen, cv2.COLOR_BGR2GRAY)  # แปลงเป็นภาพขาวดำเพื่อประมวลผล OCR ได้ง่ายขึ้น

        # ใช้ OCR เพื่อแปลงภาพเป็นข้อความ
        text = pytesseract.image_to_string(screen_gray, lang='eng')
        
        quantity_matches = re.search(r"Quantity\s+(\d+)", text)
        scanned_matches = re.search(r"Scanned\s+(\d+)", text)
        barcode_matches = re.search(r"LPN:\s*(\d+)", text)
        # upc_matches = re.search(r'UPC/EAN\(GTIN\)\s+Style.*?\n(\d+)', text, re.DOTALL)
        
        quantity = int(quantity_matches.group(1)) if quantity_matches else None
        scanned = int(scanned_matches.group(1)) if scanned_matches else None
        barcode = str(barcode_matches.group(1)) if barcode_matches else None
        # upc = int(upc_matches.group(1)) if upc_matches else None
        
        # print([quantity, scanned, barcode])
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
                print(row)
                cursor.execute("SELECT COUNT(*) FROM Tx_SP_ScanPackDet WHERE Barcode = ?", (barcode,))
                result = cursor.fetchone()
                if result[0] == 0:
                    # ถ้าไม่มีข้อมูลที่ซ้ำกัน, ทำการ INSERT
                    cursor.execute('''INSERT INTO Tx_SP_ScanPackDet (Type, Barcode, TxnId, PackTypeId, JobNo, JobDetId, PLTxnId, PPTxnId, PrePackTxnId, ColorId, SizeId, InseamId, TargetQty, ActualQty, Qty,
                    SPDate, PackListNo, CartonNo, CreateUser, CreateDateTime)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, GETDATE(), ?, ?, 'BOT', GETDATE())''',
                    (row[6], row[7], row[5], 1, row[1], row[2], 1, row[3], row[4], row[11], row[12], row[13], row[14], row[14], row[14], row[0], row[9]))
                    # ทำการ UPDATE Status
                    cursor.execute("UPDATE Tx_SP_JobDetInstPackBarcode SET Status = ? WHERE Barcode = ? AND Status IS NULL", ('PK', barcode))
                    conn.commit()
                    
                    print('update and insert success!')
                    playsound('sounds/alert.mp3')
            else:
                print("not found in the database.")
        else:
            print('quantity: '+str(quantity)+' != scanned: '+str(scanned)+' = barcode: '+str(barcode))

        # แสดงภาพหน้าจอ (สามารถข้ามขั้นตอนนี้ได้หากไม่ต้องการการแสดงผล)
        # cv2.imshow('Screen', screen_gray)

        # รอคีย์ 'q' เพื่อออกจาก
        if cv2.waitKey(1) == ord('q'):
            break

    # ปิดการเชื่อมต่อ
    cursor.close()
    conn.close()
    cv2.destroyAllWindows()

realtime_ocr()