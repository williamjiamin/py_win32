# import our libraries
import win32com.client as win32
import pythoncom


# Define our Application Events
class ApplicationEvents:
    # define an event inside of the application, Be aware OnSheetActivate Syntax!
    def OnSheetActivate(self, *args):
        print("汇报~你已经选中了这个Sheet~")

# Define our Workbook Events
class WorkbookEvents:
    # define an event inside of the workbook, Be aware OnSheetSelectionChange Syntax!
    def OnSheetSelectionChange(self,*args):
        print(args[1].Address)

# get the instance which is activated right now

excel1 = win32.GetActiveObject("Excel.Application")

# assign our event to the Excel Object
excel1_events = win32.WithEvents(excel1, ApplicationEvents)

# Get our workbook (Remember to enter the correct current workbook name(例如：工作簿2))
excel1_workbook = excel1.Workbooks("工作簿2")

# assign our event to the workbook
excel1_workbook_events = win32.WithEvents(excel1_workbook,WorkbookEvents)

while True:
    # display the message
    pythoncom.PumpWaitingMessages()
