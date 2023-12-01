from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import requests
import json
import datetime
import os

wb= Workbook()
ws = wb.active
ws.tittle = "Kursy-walut"
