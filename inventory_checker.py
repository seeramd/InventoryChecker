from os import chdir #getcwd
import win32com.client as client
import csv

#change working directory to supplies folder
chdir("[Insert Your Directory Here]")
#print(getcwd())

class Item:
    #build each item in stock as an instance
    def __init__(self, name, category, subcategory, quantity, alert_threshold):
        self.name = name
        self.category = category
        self.subcategory = subcategory
        self.quantity = quantity
        #point at which the quantity could be considered 'low,' i.e. when we'll want to buy more
        #manually set for each item, but I'm musing about some machine learning algo where we predict when an item will run
        #low based on frequency of withdrawal. Need more data though, also probably a bit much for a 20 person office
        self.alert_threshold = alert_threshold

    #format and print details on inventory item instance
    def details(self, alert_info = False):
        
        #doing the cool python 3.6 way to do string formatting
        print(f'*{self.name}*')
        print(f"{self.category}/{self.subcategory}")

        if self.quantity == 1:
            x = 'is'
        else:
            x = 'are'
        
        print(f"There {x} {self.quantity} left in stock")

        if alert_info:
            print(f"When stock drops below {self.alert_threshold}, you will receive an alert")

    #check if item quantity is at or below alert threshold
    def is_low(self):
        return (self.quantity <= self.alert_threshold)


with open("Supply Inventory.csv",'r') as file:
    reader = csv.reader(file)
    next(reader)

    #store Item class instances into a list
    #the index numbers correspond to the locations of the corresponding data in the original csv
    item_list = [Item(row[0],row[5],row[6],int(row[2]),int(row[3])) for row in reader]

#create a list containing the items that are low on stock using the class method is_low()
#This could probably be done in the previous comprehension (if row[2] <= row[3]) but I might do something with the full class,
#also it's clearer probably
low_items = [item for item in item_list if item.is_low()]

#for item in low_items:
#    print(item.name)

if len(low_items) > 0:

    email_text = "The following inventory items are low on stock:\n\n"
    for item in low_items:
        email_text += f"{item.name} ({item.quantity} left)\n"

    #Send alert email in Outlook client with pywin32
    outlook = client.Dispatch("Outlook.Application")

    message = outlook.CreateItem(0)
    message.Display()
    message.To = "[Insert Destination Email Here]"
    message.Subject = "Low Inventory Alert"

    message.Body = email_text
