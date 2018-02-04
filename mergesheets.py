import csv
import xlrd
import pandas,os
""" merge the mulitiple sheet by there multiple names """
arr_txt = [x for x in os.listdir('.') if x.endswith(".xlsx")]
nooffiles=len(arr_txt)-1
print(arr_txt)
for i in arr_txt:
    print(i)
    unwanted,newname=i.split("Report_")
    print(newname)
    os.rename(i,newname)



arr_txt2= [x for x in os.listdir('.') if x.endswith(".xlsx")]
print(arr_txt2)
'''list of sheets  '''
path=arr_txt2[0]
xls = pandas.ExcelFile(path)
sheets = xls.sheet_names
sheetsnames=[]
for i in sheets:
    sheetsnames.append(i)

for i in sheetsnames:
    print(str(i)+"\n")

print(len(sheetsnames))
'''gggh'''


waste=open("waste.txt",'w')
os.mkdir("op")

for l in sheetsnames:
    your_csv_file = open('op\\'+str(l)+'.csv', 'wb')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
    fileno=0
    for i in arr_txt2:
        wb = xlrd.open_workbook(i)
        print(i)
        try :
            sh = wb.sheet_by_name(l)
            print(l)
            for rownum in xrange(sh.nrows):
                if(fileno!=0 and rownum==0):#exceed the first row of the other than first file
                    rownum=rownum+1
                    waste.writelines(str(sh.row_values(rownum)))
                    #print(rownum)
                else :
                    wr.writerow(sh.row_values(rownum))
                    #print(rownum)


            print(rownum)
            fileno+=1
        except Exception as e:
             print("error {}".format(e)+" sheet name is diffrent or no there ")



    your_csv_file.close()
    print(l+"completed")
