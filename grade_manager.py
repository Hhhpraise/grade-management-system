from tkinter import *
from tkinter import ttk
from tkinter.filedialog import *
from tkinter.messagebox import *

# Global variables
FileName = ''
scaleOption = [0.1, 0.2, 0.7]  # weights for regular, midterm, final scores

# Create main window
root = Tk()
root.title('Grade Management System')
style = ttk.Style()
style.configure("Treeview.Heading", foreground='#0000FF')

# Main frame with Treeview and Scrollbar
fmain = LabelFrame(text='Data:')
fmain.pack(anchor=W, expand=YES, fill=BOTH)

# Define columns
columns = ('ID', 'Name', 'Regular', 'Midterm', 'Final', 'Total')
treeview = ttk.Treeview(fmain, show="headings", columns=columns)

# Configure columns
for c in columns:
    treeview.column(c, width=100, anchor='center')
    treeview.heading(c, text=c)

treeview.pack(side=LEFT, expand=YES, fill=BOTH)

# Scrollbar
sc = Scrollbar(fmain, orient=VERTICAL)
sc.pack(side=RIGHT, fill=Y)
sc.config(command=treeview.yview)
treeview.config(yscrollcommand=sc.set)


def treeview_sort_column(table, col, reverse):
    """Sort table by column when header is clicked"""
    ldata = [(table.set(k, col), k) for k in table.get_children()]
    score_columns = ('Regular', 'Midterm', 'Final', 'Total')

    def key_func(a):
        try:
            return float(a[0])  # Convert to float for numeric sorting
        except ValueError:
            return a[0]  # Fallback to string sorting

    if col in score_columns:
        ldata.sort(key=key_func, reverse=reverse)
    else:
        ldata.sort(reverse=reverse)

    for index, (val, k) in enumerate(ldata):
        table.move(k, '', index)

    table.heading(col, command=lambda: treeview_sort_column(table, col, not reverse))


# Bind sort function to column headers
for col in columns:
    treeview.heading(col, text=col,
                     command=lambda _col=col: treeview_sort_column(treeview, _col, False))


def edit_cell(event):
    """Edit cell on double-click"""
    if len(treeview.selection()) < 1:
        return

    thisitem = treeview.selection()[0]
    item_text = treeview.item(thisitem, "values")
    column = treeview.identify_column(event.x)
    row = treeview.identify_row(event.y)
    cn = int(str(column).replace('#', ''))
    rn = int(str(row).replace('I', ''), 16)

    ftool = Frame(root)

    def exitedit(event):
        ftool.destroy()

    ftool.bind('<FocusOut>', exitedit)
    ftool.place(x=event.x, y=event.y)

    text = StringVar()
    entryedit = Entry(ftool, textvariable=text)
    text.set(item_text[cn - 1])
    entryedit.pack(side=LEFT)

    def saveedit():
        treeview.set(thisitem, column=column, value=text.get())
        ftool.destroy()

    def escedit():
        ftool.destroy()

    okb = ttk.Button(ftool, text='OK', width=4, command=saveedit)
    okb.pack(side=LEFT)
    escb = ttk.Button(ftool, text='Cancel', width=4, command=escedit)
    escb.pack(side=LEFT)
    entryedit.focus_set()


treeview.bind('<Double-1>', edit_cell)


def newFile(mfile=None):
    """Clear table and current file"""
    global FileName
    items = treeview.get_children()

    if len(items) > 0:
        y = askyesno('Confirm', 'Save current data before clearing?')
        if y:
            saveRecords(mfile)

    FileName = ''
    fmain.config(text='Data:')

    for n in treeview.get_children():
        treeview.delete(n)
    treeview.update()


def openFile(fn=None):
    """Open file and load data into table"""
    if fn is not None:
        fname = fn
    else:
        fname = askopenfilename()
        if fname == '':
            return

    global FileName
    FileName = fname
    data = []

    try:
        with open(fname) as f:
            for r in f:
                r = r.strip()
                d = r.split(',')
                d.append('')  # Add empty field for Total
                data.append(d)

        data[0][5] = 'Total'  # Header for Total column
        rows = len(data)
        cols = 6

        for n in treeview.get_children():
            treeview.delete(n)

        for i in range(1, rows):
            treeview.insert('', i, text=f'r{i}', values=data[i][:cols])

        fmain.config(text=f'Data: {fname}')
    except Exception as e:
        msg = f'{fname}\nFile format error! Expected format:\n' \
              'ID,Name,Regular,Midterm,Final\n' \
              '19001,John,94,78,89\n' \
              '...'
        showerror('File Error', msg)


def saveRecords(mfile=None):
    """Save data to current file"""
    if FileName != '':
        with open(FileName, 'w') as f:
            print('ID,Name,Regular,Midterm,Final,Total', file=f)
            for n in treeview.get_children():
                row = treeview.item(n, "values")
                print(','.join(row), file=f)
        showinfo('Save', f'Data saved to: {FileName}')
    else:
        saveAsNew()


def saveAsNew():
    """Save data to new file"""
    global FileName
    fnew = asksaveasfilename(filetypes=[('CSV files', '.csv')],
                             initialfile=FileName)
    if fnew == '':
        return

    with open(fnew, 'w') as f:
        print('ID,Name,Regular,Midterm,Final,Total', file=f)
        for n in treeview.get_children():
            row = treeview.item(n, "values")
            print(','.join(row), file=f)

    fmain.config(text=f'Data: {fnew}')
    FileName = fnew
    showinfo('Save As', f'Data saved to: {fnew}')


def delRecord():
    """Delete selected record"""
    if len(treeview.selection()) < 1:
        return

    y = askyesno('Confirm', 'Delete selected record?')
    if y:
        thisitem = treeview.selection()[0]
        treeview.delete(thisitem)


def newRecord():
    """Add new empty record"""
    columns = ('New', '', '', '', '', '')
    n = treeview.insert('', 0, values=columns)
    treeview.see(n)
    treeview.selection_set(n)


def insertRecord():
    """Insert new empty record at selected position"""
    columns = ('New', '', '', '', '', '')

    if len(treeview.selection()) < 1:
        rn = 0
    else:
        row = treeview.selection()[0]
        rn = int(str(row).replace('I', ''), 16) - 1
        if rn < 0:
            rn = 0

    n = treeview.insert('', rn, values=columns)
    treeview.selection_set(n)
    treeview.see(n)


def calTotal():
    """Calculate total score based on weights"""
    for n in treeview.get_children():
        row = list(treeview.item(n, "values"))
        try:
            regular = float(row[2]) if row[2] else 0
            midterm = float(row[3]) if row[3] else 0
            final = float(row[4]) if row[4] else 0
            total = regular * scaleOption[0] + midterm * scaleOption[1] + final * scaleOption[2]
            treeview.set(n, column=5, value=str(round(total)))
        except (ValueError, IndexError):
            pass


def setOption():
    """Set score weights"""
    global scaleOption

    ftool = LabelFrame(root, text='Set Score Weights')

    s0 = DoubleVar(value=scaleOption[0])
    s1 = DoubleVar(value=scaleOption[1])
    s2 = DoubleVar(value=scaleOption[2])

    Label(ftool, text='Regular:').grid(row=1, column=1)
    Entry(ftool, textvariable=s0).grid(row=1, column=2)

    Label(ftool, text='Midterm:').grid(row=2, column=1)
    Entry(ftool, textvariable=s1).grid(row=2, column=2)

    Label(ftool, text='Final:').grid(row=3, column=1)
    Entry(ftool, textvariable=s2).grid(row=3, column=2)

    status = Label(ftool, text='')
    status.grid(row=4, column=1, columnspan=2)

    def saveedit():
        try:
            a = s0.get()
            b = s1.get()
            c = s2.get()

            if not (0 <= a <= 1 and 0 <= b <= 1 and 0 <= c <= 1) or (a + b + c) != 1:
                status.config(text='Error: Weights must be 0-1 and sum to 1!', fg='red')
            else:
                scaleOption[0] = a
                scaleOption[1] = b
                scaleOption[2] = c
                ftool.destroy()
        except:
            status.config(text='Error: Invalid input!', fg='red')

    def escedit():
        ftool.destroy()

    def exitedit(event):
        ftool.destroy()

    ftool.bind('<FocusOut>', exitedit)

    ttk.Button(ftool, text='OK', command=saveedit).grid(row=3, column=3)
    ttk.Button(ftool, text='Cancel', command=escedit).grid(row=4, column=3)

    ftool.place(x=100, y=20)
    ftool.focus_set()


def editMgx():
    """Show edit instructions"""
    showinfo('Edit Cell', 'Double-click a cell to edit it')


def sortMgx():
    """Show sort instructions"""
    showinfo('Sort Table', 'Click column headers to sort')


# Create menu system
menuroot = Menu(root)
root.config(menu=menuroot)

# File menu
mfile = Menu(menuroot, tearoff=0)
menuroot.add_cascade(label='File', menu=mfile)
mfile.add_command(label='New', command=lambda: newFile(mfile))
mfile.add_command(label='Open...', command=openFile)
mfile.add_separator()
mfile.add_command(label='Save', command=lambda: saveRecords(mfile))
mfile.add_command(label='Save As...', command=saveAsNew)
mfile.add_separator()
mfile.add_command(label='Exit', command=root.destroy)

# Management menu
medit = Menu(menuroot, tearoff=0)
menuroot.add_cascade(label='Manage', menu=medit)
medit.add_command(label='Add Record', command=newRecord)
medit.add_command(label='Delete Record', command=delRecord)
medit.add_command(label='Edit Help', command=editMgx)
medit.add_separator()
medit.add_command(label='Sort Help', command=sortMgx)

# Calculation menu
medit1 = Menu(menuroot, tearoff=0)
menuroot.add_cascade(label='Calculate', menu=medit1)
medit1.add_command(label='Set Score Weights', command=setOption)
medit1.add_command(label='Calculate Totals', command=calTotal)

root.mainloop()