
@echo off
if not exist ".venv\" py -3.13 -m venv .venv
call .venv\Scripts\activate
pip install -r requirements.txt
pause