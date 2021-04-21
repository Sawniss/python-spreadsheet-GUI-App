import gspread
from tkinter import *
from tkinter.font import Font
from tkinter import messagebox
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)

client = gspread.authorize(creds)

sheet = client.open("demosheet").worksheet("1")
# sheet = client.open("demosheet").sheet1

root = Tk()
root.title('Spreadsheet')
root.geometry("400x600")

# Define fonts
headings = Font(
    family="Impact",
    size=22,
    weight="bold"
)

normal = Font(
    family="Helvetica",
    size=19,
    weight="normal"
)

def update(*args):
    x = enterInfo.get()
    x = x.split()
    date = int(x[0])
    price = int(x[1])
    print(x)

    nameSelected = clicked.get()

    priceStr = str(price)
    dateStr = str(date)

    outputText = Label(root, text="Inserted Rs. " + priceStr + " to " + nameSelected + " on " + dateStr, font="normal")
    outputText.pack()

    name = 0
    if nameSelected == "Liam":
        name = 2
    elif nameSelected == "Noah":
        name = 3
    elif nameSelected == "Oliver":
        name = 4
    elif nameSelected == "William":
        name = 5
    elif nameSelected == "Elijah":
        name = 6
    elif nameSelected == "James":
        name = 7
    elif nameSelected == "Benjamin":
        name = 8
    elif nameSelected == "Lucas":
        name = 9

    cellstrng = ["XX", "A", "B", "C", "D", "E", "F", "G", "H", "I"]
    cellname = ''

    j = 1
    while j < 27:
        cellname = cellstrng[name]
        j += 1

    cellnumb = str(date + 1)
    commentcell = cellname + cellnumb;

    newDate = date + 1
    getCellValue = sheet.cell(newDate, name).value

    if getCellValue == None:
            sheet.update_cell(date + 1, name, price)
            enterInfo.delete(0, "end")
    else:
        newCellValue = getCellValue.replace(',','')
        sheet.update_cell(date + 1, name, int(price) + int(newCellValue))
        sheet.insert_note(commentcell, "Added " + str(price) + " & " + str(newCellValue))
        enterInfo.delete(0, "end")

    print('The value of row 5 and column 3 is :  ' + str(getCellValue))
    print(type(getCellValue))

def showData(*args):
    nameSelected = clicked.get();
    name = 0
    if nameSelected == "Liam":
        name = 2
    elif nameSelected == "Noah":
        name = 3
    elif nameSelected == "Oliver":
        name = 4
    elif nameSelected == "William":
        name = 5
    elif nameSelected == "Elijah":
        name = 6
    elif nameSelected == "James":
        name = 7
    elif nameSelected == "Benjamin":
        name = 8
    elif nameSelected == "Lucas":
        name = 9

    colDate = sheet.col_values(1)
    colValue = sheet.col_values(name)

        # outputData = Label(root, text=x)
        # outputData.pack()


# Drop Down Boxes

clicked = StringVar()
clicked.set("Select")
drop = OptionMenu(root, clicked,
                        "Liam",
                        "Noah",
                        "Oliver",
                        "William",
                        "Elijah",
                        "James",
                        "Benjamin",
                        "Lucas")
drop.pack(pady=10)

# Enter Label
label = Label(root, text="Enter", font="headings")
label.pack(pady=5)

# Date and Amount Input
enterInfo = Entry(root, width=30, text="Enter Date")

# Hit Enter to Submit
enterInfo.bind("<Return>", update)
enterInfo.pack(pady=10)

myButton = Button(root, text="Submit", command=update)
myButton.bind('<Return>', update)
myButton.pack(pady=20)

data = sheet.get_all_records()
row = sheet.row_values(2)
col = sheet.col_values(2)

root.mainloop()
