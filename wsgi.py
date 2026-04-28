import sys
import os

# Make sure Python can find app.py in this same folder
project_home = os.path.dirname(os.path.abspath(__file__))
if project_home not in sys.path:
    sys.path.insert(0, project_home)

from app import app, init_db

# Create tables and run migrations on first request
init_db()

# PythonAnywhere expects an object named 'application'
application = app
