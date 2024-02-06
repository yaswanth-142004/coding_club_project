import openpyxl as px
path = "dbms.xlsx"
wb_obj = px.load_workbook(path)

sheet_obj = wb_obj.active
row = sheet_obj.max_row
col = sheet_obj.max_column
print("Total rows:",row)
print("total columns:",col)
print("enter the contact name to search:")
NAME = str(input())
NAME = NAME.lower()
name = sheet_obj.cell(row=1,column=1)
Mobile_number = sheet_obj.cell(row=1,column=2)
Email_id = sheet_obj.cell(row=1,column=3)
Registration_number = sheet_obj.cell(row=1,column=4)
Course_enrolled = sheet_obj.cell(row=1,column=5)
my_dict ={
    "1":name,
    "2":Mobile_number,
    "3":Email_id,
    "4":Registration_number,
    "5":Course_enrolled
}



for i in range(1,row+1):
    cell_obj = sheet_obj.cell(row=i,column=1)
    if cell_obj.value.lower() == NAME:
        for j in range(1,col+1):
            cell_obj = sheet_obj.cell(row=i,column=j)
            print(my_dict[str(j)].value,end=" : ")
            print(cell_obj.value,end=" \n")
            print()