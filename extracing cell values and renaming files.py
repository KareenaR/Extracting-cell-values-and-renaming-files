import os
import xlrd

#Specify the path of your folder
path = "C:/Users/xxxx/xxxx"

files = os.listdir(path)
print(files)

#I'm extracting the cell value B7 from Sheet 0 and assiging it to a new variable named 'final'
for f in files:
    workbook = xlrd.open_workbook(path + f)
    worksheet = workbook.sheet_by_name('Sheet0')
    #Extracting row and column value
    var = worksheet.cell(6, 1)

    final = var.value
    print(final)
    #Defining the original file
    src = os.path.join(path, f)
    #Defining the final file name
    dst = path + final + ".xls"
    #If there are duplicate files, this command will just print it out, otherwise the file will be renamed
    if os.path.exists(dst):
        print(dst)
    else:
        new_filename = os.rename(src, dst)
    
