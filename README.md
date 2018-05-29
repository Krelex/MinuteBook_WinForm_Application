## How to use
Download repo, start **GrafikonSatnicaTest.sln**, set **Grafikon.View** (WinForm.app) or **GrafikonSatnicaTest**(Console test app) as StartUp project and start debuging

## Description
This app is designed to speed up working process of data entry to xls. templet. Only required fields are name, surname , month and year. Based on the given input fills the xls. sheet of working hours and ignore holidays and week. Have some additional options like open xls. after creating, choosing save path if we dont like default, choosing file name and etc.

DLLs I used for creating this app business logic are :
* [Nager.Date](https://github.com/tinohager/Nager.Date) -To calculate holidays for given month and year.
* [NPOI](https://github.com/tonyqus/npoi) - For write and read xls. templet.
* [microsoft.office.interop.excel](https://msdn.microsoft.com/library/microsoft.office.interop.excel.aspx) - To open .xls after save.
