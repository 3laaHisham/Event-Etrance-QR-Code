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


# Adds header and create id & QRCode coulmns
def header_with_id(df):
    df.columns = ['name', 'email']

    df['id'] = range(1, len(df) + 1)
    df['QRCode'] = None

def retrieve_info(df):
    data = []
    for index, row in df.iterrows():
        data_row = {}
        for col_name, value in row.items():
            data_row[col_name] = value
        data.append(data_row)

    return data

def QRCode_logo():
    logo = Image.open(Logo_path)

    basewidth = 100
    # adjust image size
    wpercent = (basewidth/float(logo.size[0]))
    hsize = int((float(logo.size[1])*float(wpercent)))
    logo = logo.resize((basewidth, hsize), Image.Resampling.LANCZOS)

    return logo

def add_QRCode_text(data_row, QRcode):
    text = yaml.dump(data_row, sort_keys= False)
    QRcode.add_data(text)

def make_QRCode(QRcode):
    QRcode.make()
    QR_img = QRcode.make_image().convert('RGB')

    return QR_img

def add_QRCode_logo(QRCode_logo, QR_img):
    pos = ((QR_img.size[0] - QRCode_logo.size[0]) // 2,
       (QR_img.size[1] - QRCode_logo.size[1]) // 2)
    QR_img.paste(QRCode_logo, pos)

def save_QRCode(data_row, QR_img):
    id = data_row['id']
    qrimg_path = os.path.join(QRCode_folder_path, str(id)) + '.png'
    QR_img.save(qrimg_path)

    return qrimg_path

# Creates a unique QR Code for each attendee and return path
def make_QR_images(data_row, QRCode_logo):

    QRcode = qrcode.QRCode(
        error_correction=qrcode.constants.ERROR_CORRECT_H)

    add_QRCode_text(data_row, QRcode)

    QR_img = make_QRCode(QRcode)

    add_QRCode_logo(QRCode_logo, QR_img)

    qrimg_path = save_QRCode(data_row, QR_img)

    return qrimg_path

def generate_qrcode(df, QRcode_logo):
    data = retrieve_info(df)

    qr_paths =[make_QR_images(data_row, QRcode_logo) for data_row in data]

    df['QRCode'] = qr_paths


if __name__ == '__main__':

    file_name = input('Enter your excel file name without the extension & Make sure it is in the same directory as the script \n')
    df = pd.read_excel(file_name + '.xlsx', header=None, index_col=False)

    header_with_id(df)

    try:
        os.mkdir('QR Code Images')
    except:
        pass
    
    QRcode_logo = QRCode_logo()
    generate_qrcode(df, QRcode_logo)

    df.to_excel(file_name + '_modified.xlsx', index=False)

    print('You have successfully generated QR Codes for every attendee.')
