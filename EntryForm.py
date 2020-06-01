from tkinter import*

def show_values(value):
    print(value)

def make_form(root):
    label_0 = Label(root, text="MERGE PDFS", width=20, font=("bold", 20))
    label_0.place(x=90, y=53)

    label_1 = Label(root, text="Input Folder Location", width=20, font=("bold", 10), anchor=W, justify=LEFT)
    label_1.place(x=80, y=130)
    inputFolder = Entry(root)
    inputFolder.place(x=240, y=130)

    label_2 = Label(root, text="Output Folder Location", width=20, font=("bold", 10), anchor=W, justify=LEFT)
    label_2.place(x=80, y=180)
    outputFolder = Entry(root)
    outputFolder.place(x=240, y=180)

    label_3 = Label(root, text="Merged File Name", width=20, font=("bold", 10), anchor=W, justify=LEFT)
    label_3.place(x=80, y=230)
    fileName = Entry(root)
    fileName.place(x=240, y=230)


if __name__=="__main__":
    root = Tk()
    root.geometry('500x400')
    root.title("Merge PDF Utility")
    ents= make_form(root)
    # it is use for display the registration form on the window
    button= Button(root, text='Merge PDFS', width=20, bg='brown', fg='white', command=show_values("Hello World"))
    button.place(x=180, y=300)
    root.mainloop()