@echo off
cd /d "%~dp0"
title SCADA Report Server
echo Starting Server...
python server.py
pause
