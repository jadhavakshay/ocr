from flask import Flask
from PIL import Image
#from flask_restful import Resource, Api
from flask import request
#from flask_api import FlaskAPI

import io
import re
import os
os.chdir(os.getcwd())
import json
import numpy as np
import pandas as pd
#import cv2

import datetime as dt
import six
import shutil
from glob import glob
from google.cloud import vision

# from google.cloud import language
# from google.cloud.language import enums
# from google.cloud.language import types

import requests
import base64
from pdf2jpg import pdf2jpg
###
import pandas as pd
from pandas.compat import StringIO
import numpy as np
from PIL import Image
import pytesseract
from pytesseract import image_to_string, image_to_osd
import base64
import io
import json
###

#from file_encrypter import image_64_encode
#from Invoice_reader_format import json_output

# Set Google API authentication and set folder where images are stored
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'Banking-326c0d0e12c1.json'
client = vision.ImageAnnotatorClient()


app = Flask(__name__)

@app.route('/process_scan', methods=['POST'])

def process_scan():

    pdf_64_encode = request.get_data()
    image_64_encode= pdf64_to_img64(pdf_64_encode)
    ##
    content = base64.decodebytes(image_64_encode)

    dataBytesIO = io.BytesIO(content)
    im = Image.open(dataBytesIO)
    data_crop = to_create_datacrop_main(im)
    #data_crop.to_json('Result1_table.json', orient='records', lines=False)

    ##




    ##google
    response = client.document_text_detection({'content': content})  # [1]
    texts = response.text_annotations
    rendered_text = texts[np.argmax([len(t.description) for t in texts])].description.split('\n')

    corpus = [' '.join(rendered_text)]

    Data = information_extract(rendered_text)

    Data = pd.DataFrame(Data)

    Data = Data.transpose()
    Data

    Invoice_Description = Data.to_json(orient='records', lines=False)
    Invoice_Description = json.loads(Invoice_Description)
    Invoice_Description = json.dumps(Invoice_Description)

    Invoice_details = data_crop.to_json(orient='records', lines=False)
    Invoice_details = json.loads(Invoice_details)
    Invoice_details = json.dumps(Invoice_details)

    import re
    p = re.compile(r"[.*^}]]")
    Invoice_Description = re.sub(p, '', Invoice_Description)

    Invoice_details = Invoice_details + '}]'
    result = Invoice_Description + "," + str("\"LineItems\"") + ":" + Invoice_details

    output = json.loads(result)

    with open('Final_Result.json', 'w') as outfile:
        json.dump(output, outfile)

    with open('Final_Result.json') as json_data:
        jsonresult = json_data.read()

    print(jsonresult)

    #os.remove('Result1_table.json')
    #os.remove('Result1.json')
    os.remove('sample.pdf')

    shutil.rmtree('data\\')

    return jsonresult
###

''' Functions '''


def pdf64_to_img64(pdf_read):
    with open("sample.pdf", "wb") as f:
        f.write(base64.decodebytes(pdf_read))

    inputpath = r"sample.pdf"
    outputpath = r"data"
    pdf2jpg.convert_pdf2jpg(inputpath, outputpath, pages="ALL")

    image_path = glob(os.path.join('data/sample.pdf', '*.jpg'))
    image_path = image_path[0]

    image = open(image_path, 'rb')
    image_read = image.read()
    image_64_encode = base64.encodebytes(image_read)
    return image_64_encode

def rendered(text):
    out = []
    buff = []
    a = text
    for c in a:
        if c == '\n':

            out.append(''.join(buff))
            buff = []

        else:
            buff.append(c)
    else:
        if buff:
            out.append(''.join(buff))
    return out


def rendered_list(rendered_text_crop):
    out = []
    buff = []
    a = rendered_text_crop
    for c in a:
        if c == ' ':

            out.append(''.join(buff))
            buff = []

        else:
            buff.append(c)
    else:
        if buff:
            out.append(''.join(buff))
    return out


def crop(im, coords, saved_location):
    image_obj = im
    # image_obj = Image.open(image_path)
    cropped_image = image_obj.crop(coords)
    # cropped_image.save(saved_location)
    # cropped_image.show()
    return cropped_image


def to_create_datacrop_main(im):
    img = crop(im, (170, 1450, 2350, 2790), 'cropped.jpg')  ## used
    text_crop = pytesseract.image_to_string(img, lang='eng')

    rendered_text_crop = rendered(text_crop)
    data = {}
    for i in range(0, len(rendered_text_crop)):
        if len(rendered_list(rendered_text_crop[i])) > 3:  # and len(rendered_list(rendered_text[i]))< 6:
            data['a%d' % i] = rendered_list(rendered_text_crop[i])

    mylist = []
    for key, value in data.items():
        mylist.append(len(value))
        max_occurance = max(mylist, key=mylist.count)
    for i in range(0, 5):
        for j in range(0, len(data)):
            if len(data[list(data)[j]]) != max_occurance:
                maxlist = data[list(data)[j]]
                maxlist[0:2] = [''.join(maxlist[0:2])]
                data['a%d' % j] = maxlist

    data_crop = pd.DataFrame.from_dict(data)
    data_crop = data_crop.transpose()
    header = data_crop.iloc[0]
    data_crop = data_crop[1:]
    data_crop.columns = header
    data_crop = data_crop.reset_index(drop=True)
    return data_crop

###



def from_addr(rendered_text):
    regex = r"(From|From:)"

    for i in range(0, len(rendered_text)):
        matches = re.match(regex, rendered_text[i])
        if matches:
            render_a = []

            for j in range(1, 6):
                render_a.append(rendered_text[i + j])

    return render_a


def to_addr(rendered_text):
    regex = r"(To:)"

    for i in range(0, len(rendered_text)):
        matches = re.match(regex, rendered_text[i])
        if matches:
            render_b = []

            for j in range(1, 5):
                render_b.append(rendered_text[i + j])

    return render_b

def invoice(rendered_text):
    regex = r"(Invoice:)"

    for i in range(0, len(rendered_text)):
        matches = re.match(regex, rendered_text[i])
        if matches:
            render_c = []

            for j in range(1, 4):
                render_c.append(rendered_text[i + j])

    return render_c


def total(rendered_text):
    regex = r"(Total|Total:)"

    for i in range(0, len(rendered_text)):
        matches = re.match(regex, rendered_text[i])
        if matches:
            render_d = []

            for j in range(1, 2):
                render_d.append(rendered_text[i + j])

    return render_d



def information_extract(rendered_text):
    rendered = rendered_text

    to = to_addr(rendered)

    frm = from_addr(rendered)

    inv = invoice(rendered)

    tot = total(rendered)

    return pd.Series({'From': frm, 'To': to, 'Invoice_detail': inv, 'Total': tot})

if __name__ == '__main__':
    app.debug = True
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
# if __name__ == '__main__':
#     app.run(debug=True)


##
