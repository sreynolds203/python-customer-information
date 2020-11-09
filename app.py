import openpyxl

file = 'revised_master_list.xlsx'
wb = openpyxl.load_workbook(file)
sheet = wb['Customer Contact List']
writer_1 = open('missing_co.csv', 'w')
writer_2 = open('missing_names.csv', 'w')
writer_3 = open('missing_PorE.csv', 'w')
writer_4 = open('missing_ship.csv', 'w')
writer_5 = open('missing_bill.csv', 'w')
for value in sheet.iter_rows(values_only = True):
    if value[0] == None:
        writer_1.write(str(value) + '\n')
    if value[1] == None or value[2] == None:
        writer_2.write(str(value) + '\n')
    if value[3] == None or value[4] == None:
        writer_3.write(str(value) + '\n')
    if value[5] == None or value[6] == None or value[7] == None or value[8] == None:
        writer_4.write(str(value) + '\n')
    if value[9] == None or value[10] == None or value[11] == None or value[12] == None:
        writer_5.write(str(value) + '\n')
    else:
        continue
writer_1.close
writer_2.close
writer_3.close
writer_4.close
writer_5.close
print('finished')