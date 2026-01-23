@echo off
REM Questo file avvia lo script PowerShell bypassando le restrizioni di sicurezza
PowerShell.exe -ExecutionPolicy Bypass -File "%~dp0convert.ps1"