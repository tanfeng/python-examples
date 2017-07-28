# -*- coding: utf-8 -*-
"""
Created by steven at 28/07/2017
add image to excel
"""

import logging
import io
import openpyxl
from openpyxl.drawing.image import Image
import requests

logging.basicConfig(format='%(asctime)s - %(name)s:%(lineno)d - %(levelname)-7s: %(message)s', level=logging.DEBUG)
logger = logging.getLogger(__name__)


def do_process():
    """
    """
    img_url = 'http://pic.qqtn.com/up/2016-10/14762726301049405.jpg'

    wb = openpyxl.Workbook()
    ws = wb.active

    # setup row height & col width
    ws.row_dimensions[1].height = 152
    ws.column_dimensions['A'].width = 25

    # fetch image
    resp = requests.get(img_url)
    img_file = io.BytesIO(resp.content)
    img = Image(img_file)
    ws.add_image(img, 'A1')

    wb.save('excel-image.xlsx')

if __name__ == "__main__":
    do_process()
    logger.info("Job Done")