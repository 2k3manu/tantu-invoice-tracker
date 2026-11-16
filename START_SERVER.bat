@echo off
echo.
echo  ====================================
echo   TANTU Bamboo Crafts - Local Server
echo  ====================================
echo.
echo  Starting local web server...
echo  Open this URL in your browser:
echo.
echo    http://localhost:8080
echo.
echo  Press Ctrl+C to stop the server.
echo  ====================================
echo.
python -m http.server 8080
pause

