@echo off
chcp 65001 > nul
echo Running docx_replace_v2_debug.ps1...
powershell -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0docx_replace_v2_debug.ps1'"