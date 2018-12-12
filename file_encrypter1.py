#from flask import Flask
#from flask_restful import Resource, Api
#from flask_api import FlaskAPI
#from Invoice_reader_format import json_output
import requests
import json

import base64

#
import os
from pdf2jpg import pdf2jpg
from glob import glob

#pdf part

# prepare headers for http request

content_type = 'PDF/pdf'
headers = {'content-type': content_type}

pdf = open('input_file.pdf', "rb").read()

pdf_64_encode =base64.encodebytes(pdf)
pdf_64_encode

#print(image_64_encode)
api_url = 'http://0.0.0.0:5000/process_scan'

# send http request with image and receive response
response = requests.post(url=api_url, data=pdf_64_encode, headers=headers)
print(response)
print(response.content)
#
