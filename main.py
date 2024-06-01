import sys
import csv
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidgetItem, QFileDialog
import generated_resource  # Ensure this import matches your actual resource file
from new_xlsx_test import export_to_excel


class MyApp(QMainWindow):
    def __init__(self):
        super(MyApp, self).__init__()
        uic.loadUi('./interface_scholar.ui', self)
        
        # Access the table widget by its object name (as defined in Qt Designer)
        self.table_widget = self.tableWidget  # Replace 'tableWidget' with the actual object name
        self.table_widget.setColumnWidth(0, 200)
        self.table_widget.setColumnWidth(2, 250)
        # Set column headers
        self.table_widget.setHorizontalHeaderLabels(['Authors', 'Year', 'Title', 'Journal', 'Volume', 'Page'])
        
        # Populate the table with data
        self.populate_table()
        
        # Add search functionality
        self.lineEdit.textChanged.connect(self.search_table)
        
        # Connect the export button to the export function
        self.exportBtn.clicked.connect(self.export_data)

    def populate_table(self):
        csv_file_path = './Research_paper_details.csv'
        
        data = []
        try:
            with open(csv_file_path, mode='r', newline='') as file:
                reader = csv.DictReader(file)
                
                for row in reader:
                    # Construct a new row with the required order and add an empty 'Citations' column
                    new_row = [
                        row['AUTHORS'] if row['AUTHORS'] != 'N/A' else '—', 
                        row['YEAR'] if row['YEAR'] != 'N/A' else '—', 
                        row['TITLE'] if row['TITLE'] != 'N/A' else '—', 
                        row['JOURNAL'] if row['JOURNAL'] != 'N/A' else '—', 
                        row['VOLUME'] if row['VOLUME'] != 'N/A' else '—', 
                        row['PAGES'] if row['PAGES'] != 'N/A' else '—'
                    ]
                    data.append(new_row)
        except Exception as e:
            print(f"Error reading CSV file: {e}")
            return
        
        # Set the number of rows
        self.table_widget.setRowCount(len(data))
        
        # Populate the rows
        for row_index, row_data in enumerate(data):
            for column_index, item in enumerate(row_data):
                self.table_widget.setItem(row_index, column_index, QTableWidgetItem(str(item)))
        
        print("Table populated successfully")  # Debugging statement

    def search_table(self):
        search_text = self.lineEdit.text().strip().lower()
        
        for row_index in range(self.table_widget.rowCount()):
            row_text = " ".join([self.table_widget.item(row_index, col).text().lower() for col in range(self.table_widget.columnCount())])
            if search_text in row_text:
                self.table_widget.setRowHidden(row_index, False)
            else:
                self.table_widget.setRowHidden(row_index, True)

    def export_data(self):
        # Open a file dialog for saving the file
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Excel Files (*.xlsx)")

        

        if file_path:
             # Check if the file path ends with ".xlsx", if not, append it
             if not file_path.endswith(".xlsx"):
                 file_path += ".xlsx"
        
        # Call the export_to_excel function
             csv_file_path = './Research_paper_details.csv'
             export_to_excel(csv_file_path, file_path)
             print(f"Data exported successfully to {file_path}")
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())
