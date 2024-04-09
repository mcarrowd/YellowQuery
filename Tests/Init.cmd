@echo off
chcp 65001

SET WORKDIR=%cd%
SET DTPATH=%WORKDIR%\1Cv8.dt
SET TESTCONN=/FO:\Tests\ib

echo Монтирую каталог...
subst o: .\..\

echo Инициализирую тестовую базу...
call runner restore --ibconnection %TESTCONN% %DTPATH%
pause
