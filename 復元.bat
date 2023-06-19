@echo off
chcp 65001
cd %~dp0
PATH=..\venv\Lib\site-packages\;%PATH%
PATH=..\venv\Scripts\;%PATH%
python main.py
pause