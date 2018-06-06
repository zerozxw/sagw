import xlwings as xw

# app=xw.App(visible=True,add_book=False)

filepath =r"F:\py workplace\aa\venv\app\2017年12维修情况.xls"
book=xw.Book(filepath)
sheet=xw.Sheet("工单列表")
sheet.activate()
str =xw.Range("A1:H2000").value

for i,maximoNO,equipmentNO,creatData,state,faultCause,repairState,time in str:
    if not i and not time:
        break
    print(i,maximoNO,equipmentNO,creatData,state,faultCause,repairState,time)










# rng=xw.Range('A:A').expand('down')
#
#
#
# nrows=rng.rows.count
#
# ncols=rng.columns.count
# print(ncols)
# print(nrows)





