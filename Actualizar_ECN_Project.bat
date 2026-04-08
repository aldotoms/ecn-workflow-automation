@echo off
title Actualizador de ECN - Flowserve
echo Iniciando actualizacion automatica...
echo.

:: 1. Entrar a la carpeta del proyecto (Con comillas para manejar espacios)
cd /d O:\11-SFM_Level_2_Planning\ECN_Project

:: 2. Ejecutar la extracción de correos
echo Paso 1: Extrayendo correos de Outlook...
python src\extractor.py

:: 3. Ejecutar el procesamiento de datos
echo Paso 2: Procesando datos y actualizando Excel...
python src\processor.py

echo.
echo ¡Proceso terminado con exito!
pause