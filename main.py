import sys
import threading
import requests
import configparser
import pandas as pd
from bs4 import BeautifulSoup
from datetime import date, datetime
from PySide6.QtWidgets import QWidget, QApplication, QMainWindow, QComboBox, QLabel, QPushButton, QGridLayout, QMessageBox

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gestor de contenido Scraping")
        self.setFixedSize(330, 75)
        self.labelInitial = QLabel("Seleccione el canal a descargar")
        
        self.listBoxChannels = QComboBox()
        self.configFile = self.read_config_ini('./config.ini')
        self.channelList = [section for section in self.configFile.sections()]
        del self.channelList[0]
        self.listBoxChannels.addItems(self.channelList)

        self.buttonDownload = QPushButton('Descargar', self)
        self.buttonDownload.clicked.connect(self.download_clicked)

        self.layout = QGridLayout()

        self.layout.addWidget(self.labelInitial, 0, 0)
        self.layout.addWidget(self.listBoxChannels, 1, 0)
        self.layout.addWidget(self.buttonDownload, 1, 3)

        widget = QWidget()
        widget.setLayout(self.layout)
        self.setCentralWidget(widget)

    def read_config_ini(self, path):
        config = configparser.ConfigParser()
        config.read(path)
        return config

    def download_clicked(self):
        channel_selected = self.listBoxChannels.currentText()

        channel_id = self.configFile[channel_selected]['id']
        channel_country = self.configFile[channel_selected]['country']
        channel_excel_name = self.configFile[channel_selected]['excel_name']

        programs = self.download_content(self.configFile, channel_country, channel_id, '-300')
        
        self.create_excel(programs, channel_excel_name)
        
    def download_content(self, config, country, channel, timezone):

        available_days = ['', 'manana', 'lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo']

        day_prog = []
        
        for day in available_days:
            if day != '':
                MAIN_URL = f"{config['settings']['url_root']}{country}{config['settings']['url_api']}{channel}/{day}/{timezone}"
            else: 
                MAIN_URL = f"{config['settings']['url_root']}{country}{config['settings']['url_api']}{channel}/{timezone}"

            # Make requests in order to get data and extract relevant content 
            page = requests.get(MAIN_URL)
            soup = BeautifulSoup(page.text, "html.parser")
            content = soup.find_all("div", {"class": "content"})
            information = soup.find_all("div", {"class": "channel-info"})
            current_date = information[0].find('span').text

            # Obtained date to a valid date for VBA        
            date_splitted_string = current_date.split(' ')
            date_splitted_string = [x for x in date_splitted_string if x.strip()]
            months = {'enero':1, 'febrero':2, 'marzo':3, 'abril':4, 'mayo':5, 'junio':6, 'julio':7, 'agosto':8, 'septiembre':9, 'octubre':10, 'noviembre':11, 'diciembre':12}
            string_d = date_splitted_string[1]+'-'+str(months[(date_splitted_string[3])])+'-'+str(date.today().year)
            converted_date = datetime.strptime(string_d, '%d-%m-%Y').date()
            converted_date = converted_date.strftime('%d-%m-%Y')
            
            for item in content:
                time = item.find("span", {"class": "time"}).text
                name = item.find('h2').text.strip()
                day_prog.append([converted_date, time, name])
            
        return day_prog

    def create_excel(self, content, excel_name):
        df = pd.DataFrame(content)
        df = df.drop_duplicates()

        df[0] = pd.to_datetime(df[0], format='%d-%m-%Y')   

        df[0] = df[0].dt.strftime('%d-%m-%Y')

        # Convert hour to datetime in order to sort it
        df[1] = pd.to_datetime(df[1])
         
        df = df.sort_values(
                by=[0, 1], ascending=True)       

        df[1] = df[1].dt.strftime('%H:%M')

        first_date_available = str(df.iloc[0,0])
        last_date_available = str(df.iloc[-1,0])

        writer = pd.ExcelWriter(f'./Grillas/{excel_name} ({first_date_available} - {last_date_available}).xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name='channels', startrow=0, index=False, header=['date', 'time', 'program'])
        writer.close()

        msg_box = QMessageBox(self)
        msg_box.setWindowTitle('Informacion')
        msg_box.setText('Se ha descargado la información del canal. Consulte la carpeta de Grillas.')
        msg_box.setIcon(QMessageBox.Information)
        msg_box.exec()


app = QApplication(sys.argv)
window = MainWindow()
window.show()

app.exec()