# PLSQL-XLSX
Create an Excel-file with PL/SQL using the XLSX files as templates

### Install
create table REPORT_TEMPLATE  
install packages:  
AS_XLSX  
AS_ZIP  
PKG_XLSX_HELPER  

### Sample
```sql
declare 
bBLOB blob;
begin
  /* create header */	
  PKG_XLSX_HELPER.CREATE_REPORT_TEMPLATE(2, 1);
  /* ... */
  :bBLOB := AS_XLSX.FINISH;
end;
```
