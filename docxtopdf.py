import sys
import os
import comtypes.client
print(os.getcwd())
wdFormatPDF = 17
print(os.listdir)


def convert_in_dir(word):
    dirspath = [x[0] for x in os.walk(os.getcwd())]
    for ue in dirspath:
        os.chdir(ue)
        
        for root, dirs, files in os.walk(os.getcwd()):
            
            for file in files:
                if file.endswith(".docx"):
                    
                    file_to_convert = os.path.join(root, file)
                    out_file = file_to_convert + ".pdf"
                    
                    doc = word.Documents.Open(file_to_convert)
                    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
                   
                    doc.Close()
                     
                    print("successfully convert " +  file_to_convert + " to pdf")
            
        dirspath = [x[0] for x in os.walk(os.getcwd())]
        if len(dirspath) > 0:
            convert_in_dir(word)

    print("successfully convert all docx files to pdf")
            
            
word = comtypes.client.CreateObject('Word.Application')
convert_in_dir(word)

word.Quit()