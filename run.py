"""
Copyright 2020 by Under the stars.


Author

Priyanshu Jain <priyanshu@pm.me>
"""

__author__ = 'priyanshu@pm.me (Priyanshu Jain)'

# Third party imports
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import pyqrcode 
from pyqrcode import QRCode 


def create_image(email, count, ticket_number):
    """
    create Image object with the input image
    
    Parameters
    ----------

    email: str
            email of the customer

    count: int
            number of tickets

    Returns
    -------
    None

    """
    image = Image.open('templates/background.jpeg')

    draw = ImageDraw.Draw(image)


    font1 = ImageFont.truetype('fonts/OpenSans-Regular.ttf', size=135)
    font2 = ImageFont.truetype('fonts/OpenSans-Regular.ttf', size=55)

    (x, y) = (295, 785)
    message = str(count)
    color = 'rgb(224, 174, 96)'
    draw.text((x, y), message, fill=color, font=font1)

    (x, y) = (240, 962)
    draw.text((x, y), email, fill=color, font=font2)

    # generate qr code
    qr_code_str = "email: {0}\ncount: {1}\nExcel row number: {2}\n".format(
        email, count, ticket_number)
    qr_code = pyqrcode.create(qr_code_str) 
    
    qr_code.png("qrcode.png", scale=5)
    qr_code = Image.open("qrcode.png")

    offset = (950, 915)
    image.paste(qr_code, offset)


    image.save('tickets/{}.png'.format(email))




def main():
    """
    Iterate over excel file and create ticket for each email and count
    """
    df = pd.read_excel('data/data.xlsx')

    for index, row in df.iterrows():
        email = row['email']
        count = row['Count']
        ticket_number = row['ticket_no']
        create_image(email, count, ticket_number)


if __name__ == "__main__":
    main()
