import csv
import xlrd
import pandas,os


arr_txt2= [x for x in os.listdir('.') if x.endswith(".xlsx")]
print(arr_txt2)

waste=open("waste.txt",'w')
os.mkdir('output')
l="report name"
your_csv_file = open('output\\'+str(l)+'.csv', 'wb')
wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
fileno=0
for i in arr_txt2:
    wb = xlrd.open_workbook(i)
    print(i)
    try :
        sh = wb.sheet_by_index(0)#sheet index No. is 0(first sheet) or sh = wb.sheet_by_name(sheetname)
        #print(l)
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
