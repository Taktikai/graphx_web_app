import os
import sys
import openpyxl
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QProgressBar, QLabel
from Xlsx_Creator import create_xls



class Sorter(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        self.setGeometry(400, 400, 500, 300)
        self.setWindowTitle('GraphX_Application')
        
        self.choose_file_btn = QPushButton('Choose File', self)
        self.choose_file_btn.move(10, 10)
        self.choose_file_btn.clicked.connect(self.chooseFile)

        self.choose_file_btn_label = QLabel('                                                                                                                ', self)
        self.choose_file_btn_label.move(110, 14)
        
        self.upload_btn = QPushButton('Create XLSX', self)
        self.upload_btn.move(10, 50)
        self.upload_btn.clicked.connect(self.makeXls)

        self.open_file_btn = QPushButton('Open XLSX', self)
        self.open_file_btn.move(10, 150)
        self.open_file_btn.clicked.connect(self.openFile)
        self.open_file_btn.hide()
        
        self.pb = QProgressBar(self)
        self.pb.setGeometry(10, 120, 380, 20)
        
        self.label = QLabel('                                                    ', self)
        self.label.move(10, 90)
        
    def makeXls(self):
        if not hasattr(self, 'file_path'):
            return

        xlsx_file_path = create_xls(self.file_path)
        
        if not xlsx_file_path:
            return

        xlsx_filename = os.path.basename(xlsx_file_path)
        self.open_file_btn.setText(f'Open {xlsx_filename}')
        self.open_file_btn.show()

        self.pb.setValue(100)
        self.label.setText('XLSX created successfully!')    

    def openFile(self):
        if not hasattr(self, 'file_path'):
            return

        filename = os.path.splitext(self.file_path)[0] + '.xlsx'

        workbook = openpyxl.load_workbook(filename)
        worksheet = workbook['Összeszerelés']
        mistakes = False
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.value == '?':
                    mistakes = True
                    break
            if mistakes:
                break

        if mistakes:
            import tkinter as tk
            from tkinter import messagebox

            root = tk.Tk()
            root.withdraw()

            answer = messagebox.askyesno("Hiányok", "Hiányosság van a vázlaton\nÍgy is meg szeretné nyitni a fájlt?")
            if answer:
                os.startfile(filename)
            else:
                sys.exit()
        else:
            os.startfile(filename)
    
    def chooseFile(self):
        self.file_path, _ = QFileDialog.getOpenFileName(self, 'Choose File', '', 'XML Files (*.graphml)')
        
        #self.proxy_file_path = create_proxy(self.file_path)
        filename = os.path.basename(self.file_path)
        
        self.choose_file_btn_label.setText(filename)
    
    

if __name__ == '__main__':
    app = QApplication(sys.argv)
    sorter = Sorter()
    sorter.show()
    sys.exit(app.exec_())