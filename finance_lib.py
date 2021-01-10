"""
Program Structure:
Store everything in a json file on disk.  Load the file into memory when running the program, populate the class and define all functions and operations within the scope of the class.  Take the class parameters and save them to disk when done modifying them.
"""

class finance:
    def __init__(self, customers, vendors, employees):
        self.customers = customers
        self.vendors = vendors
        self.employees = employees



