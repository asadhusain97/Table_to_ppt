from flask import Flask

app = Flask(__name__)
app.upload_folder = 'uploads'

from app import routes

app.static_folder = 'static'