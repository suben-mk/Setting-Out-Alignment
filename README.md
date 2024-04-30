# Setting-Out Alignment
**Setting-Out Alignment** คือตำแหน่งจุดทางเรขาคณิตของแนวอุโมงค์ (Geometry Points of Alignment) ตามตัวอย่างรูป _Alignment Type and Scheme_ วิศวกรสำรวจอุโมงค์จะต้องคำนวณตรวบสอบแบบแนวอุโมงค์ (Tunnel Alignment Drawing) ก่อนที่จะนำไปใช้ในงานสำรวจ\
ผู้เขียนได้เขียนโค้ดการคำนวณตำแหน่งจุดทางเรขาคณิตของแนวอุโมงค์ทางราบ (Geometry Points of Horizontal Alignment) ผลลัพธ์ที่ได้ระยะ Chainage และพิกัด 2 มิติ (2D-Coordinate) ของ PC, PT, TS, SC, CS, ST และตำแหน่งจุดทางเรขาคณิตของแนวอุโมงค์ทางดิ่ง (Geometry Points of Vertical Alignment) ผลลัพธ์ที่ได้ระยะ Chainage และค่าระดับ (Elevation) ของ PVC, PVT

ผู้เขียนได้ขียนโค้ดสำหรับการคำนวณ Setting-Out Alignment ไว้ 2 ภาษา คือภาษา Python และภาษา VBA Excel

### Alignment Type and Scheme
![Curve Elements - Copy Layout1 (2)](https://github.com/suben-mk/Setting-Out-Alignment-for-Tunnel-Project/assets/89971741/97fdb695-1e1e-4167-a818-07a4f407ae8a)

### Sample Setting-Out Drawing
![Hor-SO](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/assets/89971741/12bf3b28-9d49-4d32-a80d-1a328cc48d20)

![Ver-SO](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/assets/89971741/6256c041-1683-4d8c-8abd-91d89fcfc60a)

## Workflow
### Python
  **_Python libraries :_** Numpy, Pandas
  1. เตรียมข้อมูล Point of Intersection (PI) ตาม Format [Import Setting-Out Alignment Data](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/blob/main/Python/Import%20Data/Import%20Setting-Out%20Alignment%20Data.xlsx)
  2. ตั้งไฟล์ Path และจุดเริ่มต้น (Beginning Point) ของ Alignmet
     
     [**Horizontal_Alignment_Rev04.py**](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/blob/main/Python/Horizontal_Alignment_Rev04.py)
     ```py
     # Path files
     Import_data_path = "Import Setting-Out Alignment Data.xlsx"
     Export_data_path = "Export Hor-Alignment.xlsx"
      
     # Input beginning point as list[Chainage, Easting, Northing]
     BEGIN_POINT = [7202.834, 662670.304, 1521355.848]
     ```
     [**Vertical_Alignment_Rev04.py**](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/blob/main/Python/Vertical_Alignment_Rev04.py)
     ```py
     # Path files
     Import_data_path = "Import Setting-Out Alignment Data.xlsx"
     Export_data_path = "Export Ver-Alignment.xlsx"
     ```
     
  4. รันไฟล์ Python
### VBA
  1. เปิดไฟล์ [**VBA - Setting Out Alignment Program  Rev.09.xlsm**](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/blob/main/VBA/VBA%20-%20Setting%20Out%20Alignment%20Program%20%20Rev.09.xlsm)
  2. เตรียมข้อมูล Point of Intersection (PI) ที่ HIP DATA Sheet และ VIP DATA Seet

  ![hor](https://github.com/suben-mk/Setting-Out-Alignment-for-Tunnel-Project/assets/89971741/2dbd86ab-b272-4e02-ac15-086666ef9ec0)

  ![ver](https://github.com/suben-mk/Setting-Out-Alignment-for-Tunnel-Project/assets/89971741/fd7f95a4-65c8-4918-a229-8d87c37c233a)
     
  3. รันโค้ดโดยการ _คลิ๊กปุ่มสีน้ำเงิน Compute! Horizontal Alignment_ ที่ HIP DATA Sheet และ _คลิ๊กปุ่มสีน้ำเงิน Compute! Vertical Alignment_ ที่ VIP DATA Sheet

## Output
### Python
  [Export Hor-Alignment.xlsx](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/blob/main/Python/Export%20Data/Export%20Hor-Alignment.xlsx)\
  [Export Ver-Alignment.xlsx](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/blob/main/Python/Export%20Data/Export%20Ver-Alignment.xlsx)
### VBA
  [**VBA - Setting Out Alignment Program  Rev.09.xlsm**](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/blob/main/VBA/VBA%20-%20Setting%20Out%20Alignment%20Program%20%20Rev.09.xlsm)
* Horizontal Alignment ที่ HOR-SETTING OUT sheet และ HOR-ARRAY sheet
* Vertical Alignment ที่ VER-SETTING OUT sheet และ VER-ARRAY sheet
