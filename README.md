# Setting-Out Alignment
Tunnel Alignment is the center line of tunnel construction project. Tunnel survey need to crosscheck the alignment before using for survey work.
So I created the coding to compute the setting-out of horizontal and vertical alingment for reducing my mistake and time. The coding was created 2 languages which're python and vba excel.

### Alignment Type and Scheme
![Curve Elements Layout1 (1)](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/assets/89971741/0d7b41b5-93b2-4be7-b441-8ee389f86ffb)

### Sample Setting-Out Drawing
![Hor-SO](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/assets/89971741/12bf3b28-9d49-4d32-a80d-1a328cc48d20)

![Ver-SO](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/assets/89971741/6256c041-1683-4d8c-8abd-91d89fcfc60a)

### Workflow
#### Python
  1. Prepare PI data as [Import Setting-Out Alignment Data](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/blob/main/Python/Import%20Data/Import%20Setting-Out%20Alignment%20Data.xlsx)
  2. Set path file and beginning point
     
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
     
  4. Run python file
#### VBA
  1. Open file [**VBA - Setting Out Alignment Program  Rev.09.xlsm**](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/blob/main/VBA/VBA%20-%20Setting%20Out%20Alignment%20Program%20%20Rev.09.xlsm)
  2. Prepare PI data at sheet ***HIP DATA*** and sheet ***VIP DATA***
  3. Run the code by hit the buttom of ***BLUE COLOR***
     
  ![hpi](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/assets/89971741/69ceed9e-ef02-45e6-8ee9-0c0195aab47b)

  ![vpi](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/assets/89971741/b2d5c7db-4d3f-4567-b4dd-c7d7b02cd53e)

### Output
[Export Hor-Alignment.xlsx](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/blob/main/Python/Export%20Data/Export%20Hor-Alignment.xlsx)\
[Export Ver-Alignment.xlsx](https://github.com/suben-mk/Setting-Out-Alignment-for-Metro-Line/blob/main/Python/Export%20Data/Export%20Ver-Alignment.xlsx)

