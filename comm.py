import pandas as pd

class comm:
    filename = ''
    file_extensions = ['.xls', '.xlsx']
    Columns = ['Name', 'Color Prints', 'b/w Prints', 'Photo Print', 'Color Xerox', 'b/w Xerox', 'Blank Sheet', 'Scan']
    Register_file = 'names.csv'
    main_dataframe = pd.DataFrame()

    def __valid_excel(self, filename):
        if any(x in filename for x in self.file_extensions):
            return True
        else:
            return False

    def __init__(self, filename):
        if self.__valid_excel(filename):
            self.filename = filename
            try:
                self.main_dataframe = pd.read_excel(filename, sheet_name='sheet1')
            except FileNotFoundError:
                print('file not found')
                print('creating excel sheet')
                names = pd.read_csv(self.Register_file)['Names'].tolist()
                names.sort()
                print(names)
                self.main_dataframe = pd.DataFrame(columns=self.Columns)
                self.main_dataframe['Name'] = names
                self.main_dataframe.to_excel(filename, sheet_name = 'sheet1')

        else:
            print('file format is not supported')
