@echo off
echo Starting IT Operations Dashboard...
echo.
call venv\Scripts\activate
echo Server running at: http://127.0.0.1:5000
echo Press Ctrl+C to stop
echo.
python app.py
pause
