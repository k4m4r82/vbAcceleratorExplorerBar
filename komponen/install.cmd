cls
echo. install komponen pendukung
pause
copy SSubTmr6.dll %systemroot%\system32
copy vbalExpBar6.ocx %systemroot%\system32
copy vbalIml6.ocx %systemroot%\system32
regsvr32 /s %systemroot%\system32\SSubTmr6.dll
regsvr32 /s %systemroot%\system32\vbalExpBar6.ocx
regsvr32 /s %systemroot%\system32\vbalIml6.ocx