from tkinter import *
from tkinter import filedialog, ttk
from ttkthemes import ThemedStyle
# from TkinterDnD2 import DND_FILES, TkinterDnD
import tkinter.messagebox
from ttkbootstrap import Style

status = True

markDistFile = open("Mark Distribution.txt", 'r')
kLevelFile = open("K-Level.txt", 'r')

markDistFile = open("Mark Distribution.txt", 'r')
kLevelFile = open("K-Level.txt", 'r')

mark_lines = markDistFile.readlines()
K_lines = kLevelFile.readlines()
markDistList = []
kDistList = []
for i in mark_lines:
    markDistList.append(i.strip('\n'))
for i in K_lines:
    kDistList.append(i.strip('\n'))

markDistFile.close()
kLevelFile.close()

LARGEFONT = ("Helvetica", 12, 'bold')

filePath = ''
markChoice = '0'
kChoice = '0'
markNumberDict = dict()
dept = ''
lower_max_k1 = 0
upper_max_k3 = 0
lower_min_k1 = 0
upper_min_k3 = 0
SMALLFONT = ("Shruti", 11)


class tkinterApp(Tk):
    def __init__(self, *args, **kwargs):
        Tk.__init__(self, *args, **kwargs)
        container = Frame(self)
        container.pack(side="top", fill="both", expand=True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}

        for F in (LoginPage, UserPage):
            frame = F(container, self)
            self.frames[F] = frame

            frame.grid(row=0, column=0, sticky='nsew')

        self.show_frame(UserPage)


    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()


class LoginPage(Frame):
    def __init__(self, parent, controller):

        Frame.__init__(self, parent)
        titleLabel = Label(self, text="QN Paper Audit System", pady=50, padx=150, font=LARGEFONT, background = '#faf3f3',foreground='#136335')
        titleLabel.grid(row=0, column=0, columnspan=3, sticky='nsew')

        loginLabel = Label(self, text="Login as: ", font=SMALLFONT, background = '#faf3f3',foreground='#136335')
        loginLabel.grid(row=1, column=1, columnspan=1, sticky='n', pady=(20, 25))

        userButton = ttk.Button(self, text="Faculty member", width=6,
                                command=lambda: controller.show_frame(UserPage))
        userButton.grid(row=2, column=1, sticky='nsew')


class UserPage(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self, parent)

        def uploadAction(self):
            filename = filedialog.askopenfilename()
            browseLb.delete(0, END)
            browseLb.insert(0, filename)

        def drop_inside_entry(event):
            browseLb.delete(0, END)
            browseLb.insert("end", event.data)
        def onEnter(e):
            submitButton['background'] = 'black'


        def submit(event):
            global status
            if status == True:
                global filePath, markChoice, kChoice
                filePath = browseLb.get()
                if (filePath == ''):
                    tkinter.messagebox.showwarning(title="Wrong Input", message="Please upload a file", )
                elif(filePath[0] == '{' and filePath[-1] == '}'):
                    filePath = filePath[1:-1]
                elif(markChoice == ''):
                    tkinter.messagebox.showwarning(title="Wrong Input", message="Please select choice of Mark Dist", )
                elif(kChoice == ''):
                    tkinter.messagebox.showwarning(title="Wrong Input", message="Please select choice of K Level Dist", )
                elif(filePath == ''):
                    tkinter.messagebox.showwarning(title="Wrong Input", message="Please upload a file", )
                elif(filePath[-4:] != 'docx'):
                    tkinter.messagebox.showwarning(title="Wrong Input", message="File must be of docx type",)
                else:
                    app.destroy()
            else:
                tkinter.messagebox.showwarning(title = "Wrong Formatting", message = "Wrong formatting of Question Paper")

            # self.destroy()
        def getValue():
            global markChoice, kChoice
            markChoice = v.get()
            kChoice = x.get()
            # print("You selected", x.get())

        def _on_mouse_wheel(event):
            my_canvas.yview_scroll(-1 * int((event.delta / 120)), "units")

        self.configure(bg = '#ffffff')
        main_frame = Frame(self)
        main_frame.pack(fill=BOTH, expand=1)

        my_canvas = Canvas(main_frame, width=600, height=700, )
        my_canvas.pack(side=LEFT, fill=BOTH, expand=1)

        # Add A Scrollbar To The Canvas
        my_scrollbar = ttk.Scrollbar(main_frame, orient=VERTICAL, command=my_canvas.yview)
        my_scrollbar.pack(side=RIGHT, fill=Y)

        # Configure The Canvas
        my_canvas.configure(yscrollcommand=my_scrollbar.set, bg = 'white')
        my_canvas.bind('<Configure>', lambda e: my_canvas.configure(scrollregion=my_canvas.bbox("all")))
        my_canvas.bind_all("<MouseWheel>", _on_mouse_wheel)

        # Create ANOTHER Frame INSIDE the Canvas
        second_frame = Frame(my_canvas, bg = 'white')

        # Add that New frame To a Window In The Canvas
        my_canvas.create_window((0, 0), window=second_frame, anchor="nw")

        browseLabel = ttk.Label(second_frame, text="Upload a file  (.docx only)",style = 'custom.TLabel')
        browseLabel.grid(row=0, column=0, padx=10, pady=(10, 5), columnspan = 3)

        browseButton = ttk.Button(second_frame, text="Browse files", command=lambda: uploadAction(self))
        browseButton.grid(row=2, column=0, pady=(10, 10), columnspan = 3)

        browseLb = ttk.Entry(second_frame)
        browseLb.grid(row=1, column=0, ipadx=50, ipady=7, padx=125, columnspan = 3)
        # browseLb.drop_target_register(DND_FILES)
        # browseLb.dnd_bind("<<Drop>>", drop_inside_entry)

        typeQuestionLabel = ttk.Label(second_frame, text = "Select choice of Mark Distribution",style = 'custom.TLabel')
        typeQuestionLabel.grid(row = 3, column = 0, columnspan = 2,padx=80, pady=(10, 15),)

        v = StringVar()
        x = StringVar()

        currRow = 4
        variables = []

        for i in range(len(markDistList)):
            markRadioButton = ttk.Radiobutton(second_frame, text = markDistList[i], variable = v, value = str(i+1), command = getValue, style = 'info.TRadiobutton')
            markRadioButton.grid(row=currRow+i, column=0, pady=10, padx=170)
        currRow += len(markDistList)

        KlLabel = ttk.Label(second_frame, text="Select choice of K-Level distribution", style = 'custom.TLabel')
        KlLabel.grid(row=currRow, column=0, columnspan=2, padx=80, pady=(20, 15))

        for i in range(len(kDistList)):
            KLradioButton = ttk.Radiobutton(second_frame, text = kDistList[i], variable =x , value = str(i+1), command = getValue)
            KLradioButton.grid(row=currRow + i + 2 + len(kDistList), column=0, pady=10, padx=170)
        currRow += len(kDistList)

        submitButton = ttk.Button(second_frame, text="Submit", width=8, style = 'success.TButton',command=lambda: submit(self))
        submitButton.grid(row=999, column=0, columnspan = 3, pady=(20, 30))


# Driver Code

def on_close():
    global status
    status = False
    app.destroy()


app = tkinterApp()

app.title("Audit System")
app.resizable(False, False)
style = Style(theme = 'minty')

style.configure('custom.TLabel',font = LARGEFONT, foreground = '#f3969a')
app.protocol("WM_DELETE_WINDOW", on_close)
style.configure('Wild.TRadiobutton', background='#f3969a',font = SMALLFONT)


app.mainloop()

def getKLevel(st):
    x = 0
    y = 0
    l1 = st.split("%")
    print(l1)
    for n, i in enumerate(l1):
        if (len(i) > 0 and n == 0):
            x = int(i[-3:])
        if (len(i) > 0 and n == 1):
            y = int(i[-3:])
    return x, y

def getMarkDict(st):

    st = st.replace(" ", "")
    l1 = st.split(",")
    # print(l1)
    print(l1)
    val = []
    ke = []
    for i in l1:
        for n, j in enumerate(i):
            if (j == 'm'):
                num1 = int(i[0:n])
                print(num1)
                val.append(num1)
            if (j == 'x'):
                num = int(i[n + 1:])
                print(num)
                ke.append(num)
    print(val)
    print(ke)
    markslist = list(zip(ke, val))
    print(markslist)
    markNumberDict = {}
    for i, j in markslist:
        print(i, j)
        markNumberDict[i] = j
    return (markNumberDict)


if(kChoice == '1'):
    lower_max_k1, upper_max_k3 = getKLevel(kDistList[0])
elif(kChoice == '2'):
    lower_max_k1, upper_max_k3 = getKLevel(kDistList[1])
elif(kChoice == '3'):
    lower_min_k1, upper_min_k3 = getKLevel(kDistList[2])
elif(kChoice == '4'):
    lower_min_k1, upper_min_k3 = getKLevel(kDistList[3])

markNumberDict = getMarkDict(markDistList[int(markChoice) - 1])
print(lower_min_k1, lower_max_k1, upper_max_k3, upper_min_k3)
