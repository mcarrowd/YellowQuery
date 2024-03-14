@echo off
chcp 65001

SET WORKDIR=%cd%
SET BUILDPATH=%WORKDIR%\Build
SET CFE=%BUILDPATH%\YellowQuery.cfe
SET TESTCONN=/FO:\Tests\ib

mkdir %BUILDPATH%
echo Собираю расширение из исходников...
call runner compileext --ibconnection %TESTCONN% %WORKDIR%\Extensions\YQ YellowQuery
echo Применяю расширение...
call runner updateext --ibconnection %TESTCONN% YellowQuery
echo Выгружаю cfe...
call runner unloadext --ibconnection %TESTCONN% %CFE% YellowQuery
echo Собираю надстройку Excel...
cscript .\Build-xlam.vbs
pause
