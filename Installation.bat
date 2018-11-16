PowerShell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Unrestricted','-File %cd%\deployscript.ps1' -Verb RunAs"
set mypath=%cd%
@echo Solution has been installed
Pause