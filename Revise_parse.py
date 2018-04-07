# Prepare Libraries
#------------------
import openpyxl


def ParseRevSheet(revision_sheet, Table_Lenght):

    print('Du hast', Table_Lenght-2 ,'Woerter in dieser Tabelle, Wie viele woerter willst du wiederholen?')
    ziel_woerter_string=input()
    ziel_woerter=int(ziel_woerter_string)
    

    #Compare Ziel_woerter and Table_length
    #-------------------------------------
    if (ziel_woerter > Table_Lenght-2):
        print('Deine Tabelle ist kurzer als deine ziel woerter, bitte eingeben neue woerter nummer')
        ziel_woerter=int(input())
        while (ziel_woerter > Table_Lenght-2):
            print('Tabelle ist noch kurzer, bitte eingeben neue nummer')
            ziel_woerter=int(input())
        
        if (ziel_woerter == Table_Lenght-2):
            Table_Lenght=ziel_woerter
            start_index=2
        elif (ziel_woerter < Table_Lenght-2 ):
            Table_Lenght=ziel_woerter
            print('from the beginning (type b) or middle(type m) or last (type l)')
            start_pointer=input()
            if (start_pointer=='m'):
                start_index=int((Table_Lenght-ziel_woerter)/2)
                print('will start from', start_index)
            elif (start_pointer=='l'):
                start_index=(Table_Lenght-ziel_woerter)
            else:
                start_index=2
    elif (ziel_woerter == Table_Lenght-2):
        Table_Lenght=ziel_woerter
        start_index=2
    elif (ziel_woerter < Table_Lenght-2 ):
        Table_Lenght=ziel_woerter
        print('from the beginning (type b) or middle(type m) or last (type l) or define your start index (type u)')
        start_pointer=input()
        if (start_pointer=='m'):
            start_index=int((Table_Lenght-ziel_woerter)/2)
            print('will start from', start_index)
        elif (start_pointer=='l'):
            start_index=(Table_Lenght-ziel_woerter)
        elif (start_pointer=='u'):
            print('Bitte eingeben Ihre index')
            start_index=int(input())
            if (ziel_woerter>Table_Lenght-start_index):
                print('Du musst ziel woerter minizieren')
        else:
            start_index=2
    else:
        print('Bitte eingeben richtig Nummer')


    return start_index, ziel_woerter

