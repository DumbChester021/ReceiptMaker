import os
import openpyxl
class Inventory:
    class Product:
        def __init__(self, name, quantity, price):
            self.name = name
            self.quantity = quantity
            self.price = price

    def __init__(self):
        self.products = {}
        if os.path.exists("inventory.xlsx"):
            workbook = openpyxl.load_workbook('inventory.xlsx')
            sheet = workbook.active
            for row in range(2, sheet.max_row + 1):
                name = sheet.cell(row=row, column=1).value
                quantity = sheet.cell(row=row, column=2).value
                price = sheet.cell(row=row, column=3).value
                self.products[name] = self.Product(name, quantity, price)
        else:
            print("Inventory file doesn't exist")
    def add(self, name, quantity, price):
        name = name.lower()
        if name in self.products:
            product = self.products[name]
            product.quantity += quantity
        else:
            product = self.Product(name, quantity, price)
            self.products[name] = product
        self.export('inventory.xlsx')
        print("Item Added Succesfully")

    def sold(self, name, quantity):
        name = name.lower()
        if name in self.products:
            product = self.products[name]
            if product.quantity < quantity:
                print("Not enough quantity in stock")
            else:
                product.quantity -= quantity
                self.export('inventory.xlsx')
                print("Item Sold Succesfully")

        else:
            print(f"{name} not found in inventory")

    def check(self, name):
        name = name.lower()
        if name in self.products:
            product = self.products[name]
            print(f"Name: {product.name}, Quantity: {product.quantity}, Price: {product.price}")
        else:
            print(f"{name} not found in inventory")

    def check_all(self):
        if not self.products:
            print("Inventory is Empty")
        else:
            for product in self.products.values():
                print(f"Name: {product.name}, Quantity: {product.quantity}, Price: {product.price}")
    def export(self, filename):
        filename = filename.lower()
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Inventory"
        sheet.cell(row=1, column=1).value = "Product Name"
        sheet.cell(row=1, column=2).value = "Quantity"
        sheet.cell(row=1, column=3).value = "Price"

        row = 2
        for product in self.products.values():
            sheet.cell(row=row, column=1).value = product.name
            sheet.cell(row=row, column=2).value = product.quantity
            sheet.cell(row=row, column=3).value = product.price
            row += 1

        workbook.save(filename)
        print("File Saved Succesfully")

inv = Inventory()
while True:
    command = input("Enter command (add, sold, check, checkall, exit): ")
    command = command.lower()
    if command == "add":
        name = input("Enter product name: ")
        quantity = int(input("Enter quantity: "))
        price = float(input("Enter price per item: "))
        inv.add(name, quantity, price)
    elif command == "sold":
        name = input("Enter product name: ")
        quantity = int(input("Enter quantity: "))
        inv.sold(name, quantity)
    elif command == "check":
        name = input("Enter product name: ")
        inv.check(name)
    elif command == "checkall":
        inv.check_all()
    elif command == "exit":
        break
    else:
        print("Invalid command")
