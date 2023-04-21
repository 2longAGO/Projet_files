import sys
import os
import pandas as p
# Le script pour créer les fichier.conf à partir du fichier excel
#    Example:
#    python conf_maker.py path\excel_file directory_for_conf_files
#
#    Dependencies:
#    Pandas
def excel_to_conf():
    if len(sys.argv) == 3:
        if(not os.path.isfile(sys.argv[1]) and not sys.argv[1].endswith((".xls",".xlt",".xlsx",".xltx"))):
            print("The given file isn't an excel spreadsheet")
            return
        if(not os.path.isdir(sys.argv[2])):
            print("The destination path doesn't exist")
            return
        tabs = p.ExcelFile(sys.argv[1]).sheet_names
        for i in range(len(tabs)):
            df = p.read_excel(sys.argv[1],i)
            if(len(df.axes[0]) > 1 and len(df.axes[1]) > 5):
                try:
                    with open(sys.argv[2]+"\\"+tabs[i]+".conf",'x') as f:
                        for r in range(1,len(df.axes[0])):
                            # column 4 to 6
                            # name[=>][[N]][description]
                            buffer = str(df.iat[r,4]).lstrip() +"=>[" + (str(int(df.iat[r,5])) if df.iat[r,5] == df.iat[r,5] else "0") + "]" + str(df.iat[r,6]).lstrip() +"\n"
                            if(buffer[0:3] != "nan"):
                                f.write(buffer)
                        print(tabs[i]+".conf: creation completed")
                except:
                    print("The file: "+tabs[i]+".conf"+" already exists")
            else:
                print(tabs[i]+".conf: The spreadsheet has no relevant data to parse")
            
    else :
        print("Incorrect number arguments")
        print("Arg 1:excel file Arg 2:conf directory")

excel_to_conf()