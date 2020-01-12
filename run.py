"""
Author

Priyanshu Jain <priyanshu@pm.me>
"""

__author__ = 'priyanshu@pm.me (Priyanshu Jain)'

# Third party imports
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import pyqrcode 
from pyqrcode import QRCode 
from vars import LOGO_IMAGE_PATH, BLACK_COLOR, GRAY_COLOR

def copy_paste_image(source_img, target_img, offset, is_source_url=False):
    """
    create Image object by pasting source image on target image using an offset
    
    Parameters
    ----------

    source_img: PIL Image
            image that need to be pasted

    target_img: PIL Image
            image that need pasted upon

    offset: tuple
            offset position on target image
    
    is_source_url: bool
            if source image is an url

    Returns
    -------
    none

    """
    if is_source_url:
        source_img = Image.open(source_img)

    target_img.paste(source_img, offset)


def gen_qr_code(email, count, ticket_number):
    # generate qr code
    qr_code_str = "email: {0}\ncount: {1}\nExcel row number: {2}\n".format(
        email, count, ticket_number)
    qr_code = pyqrcode.create(qr_code_str) 
    
    qr_code.png("tickets/qrcode.png", scale=8)


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
    try:
        image = Image.open('templates/background.jpeg')
    except FileNotFoundError:
        image = Image.new('RGB', (1000, 1000), color = 'white')

    draw = ImageDraw.Draw(image)
    
    # paste logo
    logo_offset = (100, 50)
    copy_paste_image(LOGO_IMAGE_PATH, image, logo_offset, is_source_url=True)

    font = ImageFont.truetype('fonts/OpenSans-Regular.ttf', size=32)

    draw.text((100, 300), "SEATS", fill=BLACK_COLOR, font=font)
    draw.text((100, 350), str(count), fill=GRAY_COLOR, font=font)

    (x, y) = (100, 785)
    # draw.text((x, y), email, fill=color, font=font)

    # Add QR CODE
    gen_qr_code(email, count, ticket_number)
    qr_code = Image.open("tickets/qrcode.png")
    offset = (300, 500)
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
        ticket_number = row['ticket_number']
        create_image(email, count, ticket_number)


if __name__ == "__main__":
    main()
