@echo off
rem Install 1C - Claude bridge (asks for administrator rights)
powershell -NoProfile -Command "Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"%~dp0setup.ps1\"'"
