@echo off
rem Double-click to run. Start-Onboarding.ps1 prompts for admin (UAC) itself.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Start-Onboarding.ps1"
