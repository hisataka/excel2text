@echo off
setlocal

FOR /f "DELIMS=" %%A IN ('java -jar excel2text.jar %~1 %~2 %~3') DO SET RESULT=%%A
IF "%RESULT%" == "success" (
    exit /b 0
) else (
    exit /b 1
)
