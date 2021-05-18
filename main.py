from PyPDF2 import PdfFileReader, PdfFileWriter
from os import walk
from multiprocessing import Process, freeze_support
from docx2pdf import convert
import time
import sys



def searchPDF():
    pdf = []
    for (dirpath, dirnames, filenames) in walk("./"):
        for filename in filenames:
            if(filename[-4:] == ".pdf"):
                pdf.append(filename)
    
    return pdf


def searchWord():
    word = []
    for (dirpath, dirnames, filenames) in walk("./"):
        for filename in filenames:
            if(filename[-4:] == "docx"):
                word.append(filename)
    
    return word


def convertWord(words):
    for word in words:
        convert(word)
    #convert("input.docx", "output.pdf")
    #convert("words/")


def convertWordProcess(word):
    #print(word)
    convert(word)


def convertWordMultiProc(words):
    processes = []

    for wordFile in words:
        #print(wordFile)
        #for i in range(os.cpu_count()):
        #print('registering process %d' % i)
        processes.append(Process(target=convertWordProcess, args=(wordFile,)))

    for process in processes:
	    process.start()

    for process in processes:
        process.join()


def merge_pdfs(paths, output):
    pdf_writer = PdfFileWriter()

    for path in paths:
        pdf_reader = PdfFileReader(path)
        for page in range(pdf_reader.getNumPages()):
            # Add each page to the writer object
            pdf_writer.addPage(pdf_reader.getPage(page))

    # Write out the merged PDF
    with open(output, 'wb') as out:
        pdf_writer.write(out)



if __name__ == '__main__':
    if sys.platform.startswith("win"):
        freeze_support()

    while(True):
        print("1) Convert word to pdf")
        print("2) merge pdf files")
        print("3) exit")

        try:
            opt = int(input("->"))
        except ValueError:
            print("That's not a valid choise!")
            continue
        
        
        if(opt == 1):
            words = searchWord()
            if(len(words) > 0):
                print(words)
                start_time = time.time()
                #convertWord(words)
                convertWordMultiProc(words)
                print("--- %s seconds ---" % (time.time() - start_time))
            else:
                print("No word files found")
            continue

        if(opt == 2):
            pdfs = searchPDF()
            if(len(pdfs) >= 2):
                print(pdfs)
                merge_pdfs(pdfs, output='merged.pdf')
            else:
                print("No PDF files found")
            continue

        if(opt == 3):
            sys.exit()
        else:
            print("That's not a valid choise!")
