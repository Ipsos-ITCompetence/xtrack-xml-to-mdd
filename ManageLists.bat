@ECHO OFF

ECHO ----- Create virtual environment -----
python -m venv "%mypath%venv"


ECHO ----- Activate virtual environment -----
call .\venv\Scripts\activate

ECHO ----- Install requirements.txt -----
pip install -r .\requirements.txt


ECHO ----- Start XML to MDD process -----
python.exe ManageLists.py

ECHO ----- Deactivate virtual environment-----
call deactivate

PAUSE