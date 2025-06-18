class ExcelManager:
    def __init__(self):
        self.date_directory_path = ""
        self.xl_path = ""
    
    def set_date_path(self, date_directory_path):
        self.date_directory_path = date_directory_path

    def get_date_path(self):
        return self.date_directory_path

    def set_xl_path(self, xl_path):
        self.xl_path = xl_path 
    
    def get_xl_path(self):
        return self.xl_path
