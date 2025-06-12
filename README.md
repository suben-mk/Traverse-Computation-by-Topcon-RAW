# Traverse Computation by Topcon Raw Data
โปรแกรมวงรอบ โดยจะใช้ไฟล์ Raw จากกล้อง Total Station ของ Topcon (ผู้เขียนใช้รุ่น TOPCON GT-1201) ในตัวกล้องจะมีฟังชั่นการรังวัดมุมเป็นเซ็ต (Set Collection) ผู้เขียนใช้รูปแบบการวัดเป็น BS->FS->FS->BS ทั้งหมด 6 เซ็ต และรังวัดไป-กลับ ต่อ 1 การตั้งกล้อง (Station) โปรแกรมนี้หลักๆ จะมีดังนี้
  * แปลงข้อมูล Topcon-Raw อยู่ในรูปแบบของ MicroSurvey Star*Net สำหรับนำไปคำนวณในตัวโปรแกรม MicroSurvey
  * แปลงข้อมูล Topcon-Raw แยกเป็นแต่ละตั้งกล้อง (Station) สำหรับการเลือกมุมและระยะนำไปใช้คำนวณต่อไป
  * คำนวณวงรอบปิดแบบคู่บรรจบ (ปรับแก้แบบทิศทางและระยะ) และวงรอบปิดแบบบรรจบตัวเดียว (ปรับแก้แค่ระยะ)
  * คำนวณวงรอบเปิดแบบไม่มีคู่บรรจบ 3D-Coordinate 

## Workflow
_**Programming Language :**_ VBA Excel
1. ตั้งค่ากล้อง Classes และ Set Collection ก่อนการทำงาน

  ![Uploading 2025-06-12_164016.png…]()

2. ตั้งค่า Export Raw Data (Topcon Custom TS.txt)

   ![2025-06-12_163019](https://github.com/user-attachments/assets/a50c1614-8a27-4ad2-a504-730f3d037943)
   
3. Overview

![2025-06-12_160215](https://github.com/user-attachments/assets/417b5561-45ef-42c8-b692-2c4b9256e1dd)
![2025-06-12_160252](https://github.com/user-attachments/assets/294d8912-89d6-4846-8641-bcb5a11ddef4)
![2025-06-12_160312](https://github.com/user-attachments/assets/eae24e4e-abf1-4498-82d7-c0b014bd8490)
![2025-06-12_160331](https://github.com/user-attachments/assets/bc9874f8-5b90-4025-a1b0-93473d58b144)
![2025-06-12_160404](https://github.com/user-attachments/assets/86655fad-da0f-41d1-a5a8-777d34f66586)
![2025-06-12_160411](https://github.com/user-attachments/assets/61d2ae60-5b7e-4f9a-afaf-155416a7cde5)
![2025-06-12_160422](https://github.com/user-attachments/assets/20d78619-1d1c-49cf-86d8-3f279d598329)





