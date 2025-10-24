# ZCL_EXCEL_TO_ITAB

**ZCL_EXCEL_TO_ITAB** is an original ABAP class developed by **AnPT** to simplify the process of reading data from **Excel (.xlsx)** or **CSV (.csv)** files and converting them into **Internal Tables** in SAP.

This utility is designed for developers and functional consultants who frequently need to import structured spreadsheet data into SAP applications for data migration, reporting, or integration purposes.

---

## 🚀 Features

- 📂 Supports both **Excel (.xlsx)** and **CSV (.csv)** formats  
- 🔄 Converts data directly into **internal tables**  
- ⚙️ Flexible field mapping and automatic data type detection  
- 🧩 Easy to integrate into existing ABAP programs  
- 💡 Clear and reusable class structure  
- 🪶 Lightweight – no external dependencies  

---

## 🧠 How It Works

1. Upload your Excel or CSV file (e.g. via SAP GUI, AL11 path, or application server).
2. Instantiate the class `ZCL_EXCEL_TO_ITAB`.
3. Call the method to load and convert file contents into an internal table.
4. Use the resulting data for your business logic.

---

## 💻 Example Usage

```abap
DATA: lo_excel_reader TYPE REF TO zcl_excel_to_itab,
      lt_data         TYPE STANDARD TABLE OF string.

" Create instance
CREATE OBJECT lo_excel_reader.

" Load Excel or CSV file
lo_excel_reader->load_file(
  i_file_path = 'C:\temp\data.xlsx'    " or 'AL11 path' or 'Server path'
).

" Convert to internal table
lt_data = lo_excel_reader->get_data( ).

" Display or process data
LOOP AT lt_data INTO DATA(ls_row).
  WRITE: / ls_row.
ENDLOOP.
