from datetime import datetime
from tkinter import E

def formatSaID(text: str):
    return f'{text[0:2]}-{text[2:6]}'

def formatChID(text: str):
    return f'#{text}'

def formatDate(date, joiner: str):
        return date.strftime(f'%m{joiner}%d{joiner}%Y') if isinstance(date, datetime) else f'{date.month()}{joiner}{date.day()}{joiner}{date.year()}'