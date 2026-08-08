@echo off
rem Greenwich - Forecast Studio launcher.
rem
rem Opens the Studio front end that this repository ships, by exact name. The previous
rem launcher globbed on "09*Forecast Studio Mockup*.html" and picked the most recent match.
rem The current front end has never contained the word "Mockup", so from 17 July 2026 the
rem launcher silently opened the superseded 16 July mockup and reported nothing. A launcher
rem that cannot find its target must say so, not fall back.
rem
rem Copyright Avia Solutions Limited.
setlocal
set "TARGET=%~dp0frontend\forecast_studio.html"
if not exist "%TARGET%" (
  echo Forecast Studio front end not found at:
  echo   %TARGET%
  echo.
  echo This file ships with the repository. If it is missing, the clone is incomplete:
  echo run git status, then git pull.
  pause
  exit /b 1
)
start "" "%TARGET%"
exit /b 0
