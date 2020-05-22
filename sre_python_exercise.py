# python script excercise
# pythone version used : 3.5.2
# Author : Dashrath Goswami
# Created : 21 May 2020
# Note : Havn't use panda library beause to show basic function of python here.
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import xlrd  
Enum = []
#Console prompt for options

print("=================== Please selct your option ===================")
print("1. Display the employee name and skill set who doesn't have AWS in skills.")
print("2. Display the employee name who have more than 3 years of experienced and having atleast docker and kuberenetes in skills.")
print("3. Display the employee name in descending order who is having all skill sets.")
print("4. Exit")
option = input ()

myPath= r'data.xlsx'

#Processing request for option 1

if option == "1":
   for sh in xlrd.open_workbook(myPath).sheets():  
    for row in range(1,sh.nrows):
        #for col in range(sh.ncols):
            myCell = sh.cell(row, 2)
            #print (myCell.value)
            search = (str(myCell.value)).find('AWS')
            if search == -1:
                print('-----------')
                print('Ename :', (sh.cell_value(row,1)))
                print('Skills :', sh.cell_value(row,2))
                

#Processing request for option 2
elif option == "2":
   for sh in xlrd.open_workbook(myPath).sheets():  
    for row in range(1,sh.nrows):
        #for col in range(sh.ncols):
        exp = sh.cell(row, 3)
        if exp.value > 3.0:
            myCell = sh.cell(row, 2)
            #print (myCell.value)
            search1 = (str(myCell.value)).find('Docker')
            if search1 != -1:
                search2 = (str(myCell.value)).find('Kubernetes')
                if search2 != -1:
                    print('-----------')
                    print('Ename :', (sh.cell_value(row,1)))
                


#Processing request for option 3
elif option == "3":
   for sh in xlrd.open_workbook(myPath).sheets():  
    for row in range(1,sh.nrows):
        myCell = sh.cell(row, 2)
        search1 = (str(myCell.value)).find('Docker')
        if search1 != -1:
            search2 = (str(myCell.value)).find('Kubernetes')
            if search2 != -1:
                search3 = (str(myCell.value)).find('Python')
                if search3 != -1:
                    search4 = (str(myCell.value)).find('Jenkins')
                    if search4 != -1:
                        search5 = (str(myCell.value)).find('AWS')
                        if search5 != -1:
                            temp =(sh.cell_value(row,1))
                            Enum.append(temp)
                            #print('-----------')
                            #print('Ename :', (sh.cell_value(row,1)))
   result = sorted(Enum)
   print('-----------')
   print('Ename :', *result, sep = "\n")
                
#Processing request for option 4
elif option == "4":
   quit()

else:
    print("Wrong Option!")
    quit()