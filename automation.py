import pandas as pd
import qrcode
from PIL import Image
import os
import yaml


# Sets the default path of the script
dir_path = os.path.dirname(os.path.realpath(__file__))
os.chdir(dir_path)

QRCode_folder_path = os.path.abspath('QR Code Images')
Logo_path = 'logo.png'

class QR_Generator:

    def __init__(self, file_name):
        self.df = pd.read_excel(file_name + '.xlsx', header=None, index_col=False)

        self.header_with_id()

        try:
            os.mkdir('QR Code Images')
        except:
            pass
        
        QRcode_logo = self.QRCode_logo()
        self.generate_qrcode(QRcode_logo)

        self.df.to_excel(file_name + '_modified.xlsx', index=False)

    # Adds header and create id & QRCode coulmns
    def header_with_id(self):
        self.df.columns = ['Name', 'Email']

        self.df['id'] = range(1, len(self.df) + 1)
        self.df['QR_Link'] = None
        
        new_cols = ["id", "Name", "Email", "QR_Link"]
        self.df = self.df.reindex(columns=new_cols)
    
    def QRCode_logo(self):
        logo = Image.open(Logo_path)

        basewidth = 100
        # adjust image size
        wpercent = (basewidth/float(logo.size[0]))
        hsize = int((float(logo.size[1])*float(wpercent)))
        logo = logo.resize((basewidth, hsize), Image.Resampling.LANCZOS)

        return logo

    def generate_qrcode(self, QRcode_logo):
        data = self.retrieve_info()

        qr_paths =[self.make_QR_images(data_row, QRcode_logo) for data_row in data]

        self.df['QR_Link'] = qr_paths

    def retrieve_info(self):
        data = []
        for index, row in self.df.iterrows():
            data_row = {}
            for col_name, value in row.items():
                data_row[col_name] = value
            data.append(data_row)

        return data

    # Creates a unique QR Code for each attendee and return path
    def make_QR_images(self, data_row, QRCode_logo):

        QRcode = qrcode.QRCode(
            error_correction=qrcode.constants.ERROR_CORRECT_H)

        self.add_QRCode_text(data_row, QRcode)

        QR_img = self.make_QRCode(QRcode)

        self.add_QRCode_logo(QRCode_logo, QR_img)

        qrimg_path = self.save_QRCode(data_row, QR_img)

        return qrimg_path

    def add_QRCode_text(self, data_row, QRcode):
        text = yaml.dump(data_row, sort_keys= False)
        QRcode.add_data(text)

    def make_QRCode(self, QRcode):
        QRcode.make()
        QR_img = QRcode.make_image().convert('RGB')

        return QR_img

    def add_QRCode_logo(self, QRCode_logo, QR_img):
        pos = ((QR_img.size[0] - QRCode_logo.size[0]) // 2,
        (QR_img.size[1] - QRCode_logo.size[1]) // 2)
        QR_img.paste(QRCode_logo, pos)

    def save_QRCode(self, data_row, QR_img):
        id = data_row['id']
        qrimg_path = os.path.join(QRCode_folder_path, str(id)) + '.png'

        QR_img.save(qrimg_path)

        hyperlink = f'=HYPERLINK("{qrimg_path}", "{id}_QR_Link")'
        return hyperlink

    
if __name__ == '__main__':

    file_name = input('Enter your excel file name without the extension & Make sure it is in the same directory as the script \n')
    QR_Generator(file_name)

    print('You have successfully generated QR Codes for every attendee.')
