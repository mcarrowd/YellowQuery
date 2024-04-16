@echo off
chcp 65001

SET WORKDIR=%cd%
SET BUILDPATH=%WORKDIR%\build
SET CFE=%BUILDPATH%\YellowQuery.cfe
SET BUILDCONN=/F%BUILDPATH%\ib

mkdir %BUILDPATH%
echo Инициализирую базу для сборки...
call runner init-dev
echo Собираю расширение из исходников...
call runner compileext --ibconnection %BUILDCONN% %WORKDIR%\Extensions\YQ YellowQuery
echo Применяю расширение...
call runner updateext --ibconnection %BUILDCONN% YellowQuery
echo Выгружаю cfe...
call runner unloadext --ibconnection %BUILDCONN% %CFE% YellowQuery
echo Собираю надстройку Excel...
cscript .\Build-xlam.vbs
pause
