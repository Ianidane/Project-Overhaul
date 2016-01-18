import sys
import os
global namelength
global name
global filename

def readrowsrc(filename, namelength, name):
    procloc = 0
    fixprocloc = 0
    spc = " "
    os.chdir("C:/VBA2PHP/ProIRB_Text")
    instance = open(filename,"r") 
    
    
    
    readable = instance.read()

    if "=\"SELECT" in readable:
        while readable[procloc : procloc + namelength] != name:
            procloc = procloc + 1

        #below is the same as C# substring, the number 8 stands for the length
        while readable[procloc : procloc + 8] != "=\"SELECT":
            procloc = procloc - 1
        if readable[procloc : procloc + 8] == "=\"SELECT":
            beginquery = procloc
            while readable[procloc : procloc + 1] != ";":
                procloc = procloc + 1
            endquery = procloc
            result = readable[beginquery : beginquery + (endquery - beginquery)]
            result = ' '.join(result.split())
            result = result.replace("=\"SELECT", "SELECT")
            if "FROM" in result:
                result = result.replace("FROM", "FROM")
            if "ORDERBY" in result:
                result = result.replace("ORDERBY", "ORDER BY")
            if "\"\"" in result:
                result = result.replace("\"\"", "")
            if "WHERE" in result:
                result = result.replace("WHERE", "WHERE")
            if "LIKE" in result: 
                result = result.replace("Like", "Like")
            if "\" \"" in result:
                result = result.replace("\" \"", "")


            

            procloc = 0
            while result[procloc : procloc + 1] != "[":
                procloc = procloc + 1
            procloc = procloc + 1
            beginfilename = procloc
            while result[procloc : procloc + 1] != "[":
                procloc = procloc + 1
            procloc = procloc + 1
            beginselectname = procloc
            while result[procloc: procloc + 1] != "]":
                procloc = procloc + 1
            selectname = result[beginselectname : beginselectname + (procloc - beginselectname)]
            readrowsrc.selectname = selectname

            procloc = beginfilename
            while result[procloc : procloc + 1] != "]":
                procloc = procloc + 1
            filesname = result[beginfilename : beginfilename + (procloc - beginfilename)]
            newfilename = result[beginfilename : beginfilename + (procloc - beginfilename)]
            newfilename = newfilename + "1"
            completefilename = "Table_" + filesname + "1.txt"

            if os.path.isfile(completefilename):
                result = result.replace(filesname, newfilename)
                print result
        
    return result
    
   
       
if querytype == "needprettyquery":
    result = readrowsrc(filename, namelength, name)
    selectname = readrowsrc.selectname

