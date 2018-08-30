@echo off

SET WORKDIR=%~d0%~p0

IF "%QUARTZ%"=="" SET QUARTZ=%WORKDIR%..\

rem SET QEB=%QUARTZ%
SET QEB=%WORKDIR%

SET TMP_CP=.



dir /b "%QUARTZ%\*.jar" > temp.tmp

FOR /F %%I IN (temp.tmp) DO CALL "%QEB%\addpath.bat" "%QUARTZ%\%%I"

echo "QEB: " "%QEB%"
echo "QUARTZ: " "%QUARTZ%"

dir /b "%QUARTZ%\*.jar" > temp.tmp
FOR /F %%I IN (temp.tmp) DO CALL "%QEB%\addpath.bat" "%QUARTZ%\build\%%I"

dir /b "%QUARTZ%\lib\*.jar" > temp.tmp
FOR /F %%I IN (temp.tmp) DO CALL "%QEB%\addpath.bat" "%QUARTZ%\lib\%%I"

DEL temp.tmp
