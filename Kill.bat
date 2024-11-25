@echo off
:again
taskkill /F /IM "Mara.exe"
if errorlevel=0 goto end
if errorlevel=1 goto again
%SendKeys% {Enter}
:end

