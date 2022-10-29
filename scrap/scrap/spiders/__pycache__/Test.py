
import openpyxl

links = []
wookbook = openpyxl.load_workbook("1Comp.xlsx")

worksheet = wookbook.active


for i in range(1,100):
    for col in worksheet.iter_cols(8, 8):
        links.append(col[i].value)

print (links)
         

   
