import camelot
import pandas as pd
from pdf2image import convert_from_path

file = "D:\\Dossier Dev\\SPC\\work\\2021.378-1-3_CLIE_B42.pdf"

image_file = "./"

pages = convert_from_path(file, dpi=300, first_page=1, last_page=1)
pages[0].save(image_file, "PNG")


