
# Reading Excel-Record-to-SQL-Server


อ่านไฟล์ Excel และ ทำการเลือก ข้อมูล บันทึกข้อมูลลง SQl Server


## 1.Installation

- [Dotnet SDK Version 5.0](https://dotnet.microsoft.com/en-us/download/dotnet/5.0)
- [Visual Studio Code](https://code.visualstudio.com/download)


## 2.เชื่อมต่อกับ SQl Server
เปิดไฟล์ เชื่อมต่อกับ ฐาข้อมูล

```bash
string connectionString = "Data Source=localhost\\SQLEXPRESS;Database=DB_Excel;Integrated Security=true";
SqlConnection conn = new SqlConnection(connectionString);
```
    
## 3.ทดสอบ

To deploy this project run

```bash
  dotnet run
```


## Badges

Add badges from somewhere like: [shields.io](https://shields.io/)

[![MIT License](https://img.shields.io/apm/l/atomic-design-ui.svg?)](https://github.com/tterb/atomic-design-ui/blob/master/LICENSEs)
[![GPLv3 License](https://img.shields.io/badge/License-GPL%20v3-yellow.svg)](https://opensource.org/licenses/)
[![AGPL License](https://img.shields.io/badge/license-AGPL-blue.svg)](http://www.gnu.org/licenses/agpl-3.0)

