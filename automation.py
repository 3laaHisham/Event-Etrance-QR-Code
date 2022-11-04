import pandas as pd
import qrcode
import os
import yaml


# Sets the default path of the script
dir_path = os.path.dirname(os.path.realpath(__file__))
os.chdir(dir_path)

QRCode_folder_path = os.path.abspath('QR Code Images')


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

# Creates a unique QR Code for each attendee and return path
def make_QR_images(data_row):
    text = yaml.dump(data_row, sort_keys= False)
    qr_image = qrcode.make(text)

    id = data_row['id']
    qr_image_path = os.path.join(QRCode_folder_path, str(id)) + '.png'
    qr_image.save(qr_image_path)

    return qr_image_path

def generate_qrcode(df):
    data = retrieve_info(df)

    qr_image_paths =[]
    for data_row in data:
        path = make_QR_images(data_row)
        qr_image_paths.append(path)

    df['QRCode'] = qr_image_paths


if __name__ == '__main__':

    file_name = input('Enter your excel file name without the extension & Make sure it is in the same directory as the script \n')
    df = pd.read_excel(file_name + '.xlsx', header=None, index_col=False)

    header_with_id(df)

    try:
        os.mkdir('QR Code Images')
    except:
        pass

    generate_qrcode(df)

    df.to_excel(file_name + '_modified.xlsx', index=False)
