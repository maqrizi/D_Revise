import tkinter
from tkinter import ttk
from tkinter import messagebox
import openpyxl
import xlsxwriter
from random import shuffle

import Revise_init
import Revise_ask
import Revise_result



## Create Window
##--------------
MainWindow=tkinter.Tk()
MainWindow.title('Language Safe')
MainWindow.geometry('500x600')

TabbedWindow = ttk.Notebook(MainWindow)
TabbedWindow.pack(expand="yes", fill="both")

## Application Tabs
##-----------------
RevisionWindow=ttk.Frame(TabbedWindow, height=200, width=200)
TabbedWindow.add(RevisionWindow, text='Revise Words')



##Define Objects of the Revision Tab
##----------------------------------
TopFrame = tkinter.Frame(RevisionWindow, background="white", height=50, width=50, relief="ridge", borderwidth=5)
QA_FRAME = tkinter.Frame(TopFrame, background="#6699ff", height=50, width=50, relief="ridge", borderwidth=5)
ResultsFrame = tkinter.Frame(RevisionWindow, background="white", height=50, width=50, relief="ridge", borderwidth=5)

TopFrame.pack(expand="yes", fill="both")
QA_FRAME.pack(expand="yes", fill="y")
ResultsFrame.pack(expand="yes", fill="both")

ResultTree = ttk.Treeview(ResultsFrame)

QAFrame_Title= tkinter.Label(QA_FRAME, text="Wiederholungsfragen", background="#6699ff", font=('Arial', 12,'bold'))
QAFrame_Title.grid(row=1, column=0)

LineCanvas = tkinter.Canvas(QA_FRAME, height=10, width=20, relief="flat", background="#6699ff", bd=0, highlightthickness=0)
LineCanvas.grid(row=2, column=0, sticky="news")
LineCanvas.create_line(0, 7, 800, 7, fill="black", width=2)

Question=tkinter.Label(QA_FRAME, text="Willkommen ... Drucken Enter um zu starten", relief="flat", background="#6699ff", wraplength=300)
Question.grid(row=3, column=0, sticky="w",  padx=50, pady=25)

Answer=tkinter.Entry(QA_FRAME)
Answer.grid(row=4, column=0, padx=100, pady=5)
Answer.focus()

ResultText=tkinter.Label(QA_FRAME, text=('Status: Revision in Progress'), relief="flat", background="#6699ff")
ResultText.grid(row=5, column=0, sticky="w",  padx=5, pady=5)

LineCanvas2 = tkinter.Canvas(QA_FRAME, height=10, width=20, relief="flat", background="#6699ff", bd=0, highlightthickness=0)
LineCanvas2.grid(row=6, column=0, sticky="news")
LineCanvas2.create_line(0, 7, 800, 7, fill="black", width=2)

def quit_add():
    MainWindow.destroy()
quit_button=tkinter.Button(QA_FRAME, text='Close Application', command=quit_add)
quit_button.grid(row=7, column=0, padx=5, pady=10)

##---------------------------------------------------------------


global NextQuestion
global StartPointQuestion
NextQuestion='StartPreQ'
StartPointQuestion='StartPreQ'


##Function to Display Question
##----------------------------
def InteractiveQuestion(AskUser):
    Question.configure(text=AskUser)



def ShowResult():
    a = 1

ShowRes_button=tkinter.Button(QA_FRAME, text='Show Detailed Results', command=ShowResult)
            
def ShowClearResultsButton(mode):
    if (mode == '1'):
        ShowRes_button.grid(row=6, column=0, padx=5, pady=10)
        ShowRes_button.grid()
    elif (mode == '0'):
        ShowRes_button.grid(row=6, column=0, padx=5, pady=10)
        ShowRes_button.grid_remove()
        


##Pre-Questions to prepare revision sheet
##---------------------------------------

def PreWordsQ(QInTurn):
    global CurrentQuestion
    global NextQuestion
    global DB_Category, TableLength, WordsNo, StartIndex, WordsCounter
    global revision_sheet, revision_book
    #global DB_Category, StartIndex
    print('Q in turn is ', QInTurn)
    if (QInTurn == 'StartPreQ'):
        Q_Msg='Welche Woerter willst du wiederholen? alt/neu'
        InteractiveQuestion(Q_Msg)
        NextQuestion='WhichDB'
    elif (QInTurn == 'WhichDB'):
        DB_Category = Answer.get()
        if ((DB_Category == 'alt') or (DB_Category == 'neu')):
            revision_sheet, revision_book, TableLength = Revise_init.OpenRevSheet(DB_Category)
            NextMsg = ('Du hast {} Woerter in dieser Tabelle, Wie viele woerter willst du wiederholen?'.format(TableLength-2))
            InteractiveQuestion(NextMsg)
            Answer.delete(0, 'end')
            NextQuestion='HowManyWords'
        else:
            NextMsg = 'Bitte Geben richtige Eingabe ein (alt oder neu)'
            InteractiveQuestion(NextMsg)
            Answer.delete(0, 'end')
            NextQuestion='WhichDB'
    elif (QInTurn == 'HowManyWords'):
        WordsNo = int(Answer.get())
        if (WordsNo == 0) :
            NextMsg = 'please provide number greater than 0'
            InteractiveQuestion(NextMsg)
            Answer.delete(0, 'end')
            NextQuestion='HowManyWords'
        else:    
            WordsCounter = WordsNo
            if (WordsNo > TableLength-2):
                NextMsg = 'Deine Tabelle ist kurzer als deine ziel woerter, bitte eingeben neue woerter nummer'
                InteractiveQuestion(NextMsg)
                Answer.delete(0, 'end')
                NextQuestion='HowManyWords'
            elif (WordsNo == TableLength-2):
                StartIndex = 2
                NextMsg = 'Sie Wollen alle Woerter wiederholen, Druecken Enter um die Wiederholung zu starten'
                InteractiveQuestion(NextMsg)
                Answer.delete(0, 'end')
                NextQuestion='PRE_Q_DONE'
            elif (WordsNo < TableLength-2):
                NextMsg = 'from the beginning (type b) or from end of list (type e) or define your start index (type u)'
                InteractiveQuestion(NextMsg)
                Answer.delete(0, 'end')
                NextQuestion='WhichStartRegion'
    elif (QInTurn == 'WhichStartRegion'):
        IndexRegion = Answer.get()
        if (IndexRegion == 'b'):
            StartIndex = 2
            NextMsg = 'Druecken Enter um die Wiederholung zu starten'
            InteractiveQuestion(NextMsg)
            Answer.delete(0, 'end')
            NextQuestion='PRE_Q_DONE'
        elif (IndexRegion == 'u'):
            NextMsg = 'Bitte, geben Sie Ihre Index ein'
            InteractiveQuestion(NextMsg)
            Answer.delete(0, 'end')
            NextQuestion='WhichStartIndex'
        elif (IndexRegion == 'e'):
            StartIndex = TableLength - WordsNo
            NextMsg = 'Druecken Enter um die Wiederholung zu starten'
            InteractiveQuestion(NextMsg)
            Answer.delete(0, 'end')
            NextQuestion='PRE_Q_DONE'
        else:
            NextMsg = 'Bitte, geben Sie richtige Eingabe ein (b oder u)'
            InteractiveQuestion(NextMsg)
            Answer.delete(0, 'end')
            NextQuestion='WhichStartRegion'
    elif (QInTurn == 'WhichStartIndex'):
        StartIndex = int(Answer.get())
        if (WordsNo > TableLength-StartIndex):
            NextMsg = 'Waehlen Sie andere Index aus weil die woerter nummer ist gross'
            InteractiveQuestion(NextMsg)
            Answer.delete(0, 'end')
            NextQuestion='WhichStartIndex'
        else:
            NextMsg = 'Druecken Enter um die Wiederholung zu starten'
            InteractiveQuestion(NextMsg)
            Answer.delete(0, 'end')
            NextQuestion='PRE_Q_DONE'
    CurrentQuestion=NextQuestion

#---------------------------------Python_Fn_03---------------------------------
def WordsQuestions():
    global master_book, revision_book, master_sheet, revision_sheet
    global DB_Category, TableLength, WordsNo, StartIndex
    m = 1
    n = 1

    #Ask User and Save Answer
    #------------------------
    if (WordsCounter > 0):
        Revise_ask.OperationMode = 'GUI'
        Revise_ask.RevWordsCounter = WordsCounter
        NextMsg, CellInTurn = Revise_ask.AskRevWords(revision_sheet, StartIndex, WordsNo)
        InteractiveQuestion(NextMsg)
        UserAnswer = Answer.get()
        revision_sheet.cell(row=CellInTurn-1, column=3).value=UserAnswer
        Answer.delete(0, 'end')
    else:
        UserAnswer = Answer.get()
        CellInTurn=(WordsNo+StartIndex)
        revision_sheet.cell(row=CellInTurn-1, column=3).value=UserAnswer
        Answer.delete(0, 'end')
        NextMsg = 'Sie haben alle woerter wiedergeholt'
        InteractiveQuestion(NextMsg)
        result = Revise_result.GetRevResult(revision_sheet, StartIndex, WordsNo)
        revision_book.save('ReviseWords.xlsx')
        ResultText.configure(text=('Result is {} %'.format(result)))
        ShowClearResultsButton('0')

#---------------------------------EOF_Fn_03---------------------------------
		
def ResetRevision():
    ResultText.configure(text="  ")
    ResultTree.delete(*ResultTree.get_children())
    ShowClearResultsButton('0')
    PreWordsQ("StartPreQ")

RevisionResetButton=tkinter.Button(QA_FRAME, text='Reset Revision', command=ResetRevision)
RevisionResetButton.grid(row=8, column=0, padx=5, pady=10)

def QuestionMode(interQuestion):
    global WordsNo, WordsCounter
    if (interQuestion == 'PRE_Q_DONE'):
        WordsQuestions()
        WordsCounter = WordsCounter - 1
    else:
        PreWordsQ(interQuestion)

CurrentQuestion=NextQuestion
Answer.bind('<Return>', (lambda event: QuestionMode(CurrentQuestion)))



RevisionWindow.mainloop()
