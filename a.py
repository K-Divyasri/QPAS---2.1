import tkinter
from tkinter import messagebox
def error():
    errorMessage = """Incorrect formatting of question paper. Please adhere to these guidelines.
    
1. The document has to be in the Microsoft Word 2007/2010 Document (*.docx) format.

2. The headings for each part (part A, part B) must be above the table and must NOT be in a cell. It must be in the text form.

3. The questions for each part (A,B,C) have to be in one continuous table. 

4. The or questions have to be separated by a separate cell with the word 'Or' in it.

5. For questions with subquestions the subquestions have to be given in one cell and have to be in the given format. 
     Subquestion 1  (4 marks) 
     Subquestion 2  (2 marks)
     Subquestion 3  (2 marks)
     Subquestion 4  (2 marks)
The marks have to be given in brackets with the word 'marks' after the number. 
    
6. Please choose the correct mark distribution from the options"""
    # hide main window
    root = tkinter.Tk()
    root.withdraw()
    # message box display
    messagebox.showwarning("Error", errorMessage)
try:
    from a2 import *
except:
    error()
