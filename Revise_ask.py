# Prepare Libraries
#------------------
import openpyxl

RevWordsCounter = 0

def AskRevWords(revision_sheet, StartIndex, WordsNo):

    #Ask User and Save Answer
    #------------------------
    WordsCounter = RevWordsCounter
    asked_word=revision_sheet.cell(row=(WordsNo-WordsCounter+StartIndex), column=2).value
    asked_word=str(asked_word)
    WordOrder = str(WordsNo-WordsCounter+1)
    NextMsg = (WordOrder +'. Was ist '+ asked_word)
    CellInTurn=(WordsNo-WordsCounter+StartIndex)
    return NextMsg, CellInTurn




