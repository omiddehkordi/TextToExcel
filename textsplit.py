import pandas as pd

def textsplit(textfile):
    text = open(textfile)

    #Splits file into subsections that will each become a separate Excel Sheet
    newfile = open("text1.txt", "a")
    num = "1"
    for x in text:
        if "+ " in x:
            num = num + "1"
            newfile = open("text" + num + ".txt", "x")

        newfile.write(x)




    