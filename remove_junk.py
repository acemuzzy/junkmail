__author__ = 'Murray'

import imaplib
import email
import re
import openpyxl
from openpyxl.cell import get_column_letter
import os
import yaml

WORKSHEET_NAME = "CURRENT_CONTENT"
BLACKLIST_NAME = "BLACK_LIST"

COLUMNS = [
    "SENDER", "ID"
]

class ExcelHandler:

    def __init__(self, filename):
        self.filename = filename
        self.blacklist = []
        self.blacklist_row = 2
        self.wb = None

    def get_wb(self):
        if not self.wb:
            self.wb = openpyxl.load_workbook(self.filename)
        return self.wb

    def exists(self):
        print("Checking existence")
        return os.path.exists(self.filename)

    def create_file(self):

        print "Creating file"

        self.wb = openpyxl.Workbook()
        ws = self.wb.get_active_sheet()
        ws.title = WORKSHEET_NAME

        for col, colName in enumerate(COLUMNS):
            cell = ws.cell(row=1, column=(col+1))
            cell.value = colName
            cell.font = openpyxl.styles.Font(bold=True)

        ws.column_dimensions[get_column_letter(2)].width = 50
        ws.freeze_panes = ws['B2']

        ws = self.wb.create_sheet(BLACKLIST_NAME)
        ws.column_dimensions[get_column_letter(1)].width = 50
        ws.freeze_panes = ws['B2']

        print " Done!"
        exit(0)

    def read_blacklist(self):
        print("Reading blacklist")

        wb = self.get_wb()
        ws = wb.get_sheet_by_name(BLACKLIST_NAME)
        ws.column_dimensions[get_column_letter(1)].width = 50

        while True:
            name = ws.cell(row=self.blacklist_row, column=1).value

            if name:
                if name not in self.blacklist:
                    self.blacklist.append(name)
            else:
                break

            self.blacklist_row += 1

        print("Blacklist is {}".format(self.blacklist))

    def update_blacklist(self):
        print("Updating blacklist")
        wb = self.get_wb()
        ws = wb.get_sheet_by_name(WORKSHEET_NAME)
        bs_ws = wb.get_sheet_by_name(BLACKLIST_NAME)
        read_row = 2

        while True:
            x = ws.cell(row=read_row, column=1).value
            name = ws.cell(row=read_row, column=2).value

            if name:
                if (name not in self.blacklist) and (x == "x"):
                    self.blacklist.append(name)
                    bs_ws.cell(row=self.blacklist_row, column=1).value = name
                    print("Putting {} at {}".format(name, self.blacklist_row))
                    self.blacklist_row += 1
            else:
                break

            read_row += 1

        print("Blacklist is now {}".format(self.blacklist))

    def doWork(self, senders):
        print("Outputting data")
        wb = self.get_wb()
        ws = wb.get_sheet_by_name(WORKSHEET_NAME)

        output_row = 2
        for sender in senders:
            ws.cell(row=output_row, column=2).value = sender

            if sender in self.blacklist:
                ws.cell(row=output_row, column=1).value = "x"
                print("delete {}".format(sender))
            else:
                ws.cell(row=output_row, column=1).value = ""
                print("leave {}".format(sender))

            output_col = 3
            for id in senders[sender]["ids"]:
                if sender in self.blacklist:
                    print("delete id {}".format(id))
                ws.cell(row=output_row, column=output_col).value = id
                ws.column_dimensions[get_column_letter(output_col)].width = 6
                output_col += 1

            output_row += 1

        # clear any other rows
        while True:
            if ws.cell(row=output_row, column=1).value or \
               ws.cell(row=output_row, column=2).value or \
               ws.cell(row=output_row, column=3).value:
                ws.cell(row=output_row, column=1).value = ""
                ws.cell(row=output_row, column=2).value = ""

                output_col = 3
                while True:
                    if not ws.cell(row=output_row, column=output_col).value:
                        break
                    ws.cell(row=output_row, column=output_col).value = None
                    output_col += 1

            else:
                break

            output_row += 1

        ws.column_dimensions[get_column_letter(2)].width = 50

        wb.save(self.filename)

print("Loading configuration")

with open("imap.cfg", 'r') as stream:
    try:
        doc = yaml.load(stream)
        imap_details = doc["imap_server"]

        server = imap_details["server"]
        username = imap_details["username"]
        password = imap_details["password"]

    except yaml.YAMLError as exc:
        print("YAML error: {}".format(exc))
        exit(0)
    except KeyError as exc:
        print("Missing key: {}".format(exc))
        exit(0)

print("Config loaded")

try:
    mail = imaplib.IMAP4_SSL(server)
    mail.login(username, password)

    # Out: list of "folders" aka labels in gmail.
    mail.select("inbox") # connect to inbox.

    result, data = mail.search(None, "ALL")
except imaplib.IMAP4.error as exc:
    print("IMAP error: {}".format(exc))
    exit(0)

ids = data[0] # data is a list.
id_list = ids.split() # ids is a space separated string
print("{} emails".format(len(id_list)))
ids = id_list[-50:] # get the latest

senders = {}

for id in ids:
    result, data = mail.fetch(id, "(RFC822)") # fetch the email body (RFC822) for the given ID
    for response_part in data:
        if isinstance(response_part, tuple):
            msg = email.message_from_string(response_part[1])
            sender = msg['from'].split()[-1]
            address = re.sub(r'[<>]','',sender)

            print("{},{}".format(id, address))

            if address in senders:
                senders[address]["ids"].append(id)
            else:
                senders[address] = {}
                senders[address]["ids"] = [id]

for sender in senders:
    print("{}: {}".format(sender, senders[sender]["ids"]))

write_mode = True

if write_mode:
    writer = ExcelHandler("test.xlsx")
    if not writer.exists():
        writer.create_file()
    else:
        writer.read_blacklist()
        writer.update_blacklist()

    writer.doWork(senders)