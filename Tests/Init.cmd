@echo off
chcp 65001

SET WORKDIR=%cd%
SET DTPATH=%WORKDIR%\1cv8.dt
SET CFEPATH=%WORKDIR%\..\Build\YellowQuery.cfe
SET TESTCONN=/FO:\Tests\ib

echo Монтирую каталог...
subst o: .\..\

echo Инициализирую тестовую базу...
call runner restore --ibconnection %TESTCONN% %DTPATH%
echo Включаю аутентификацию ОС...
call runner run --ibconnection %TESTCONN% --db-user "Администратор" --execute .\Externals\DataProcessors\SetOSAuthentication\SetOSAuthentication.epf
echo Загружаю расширение из файла...
call runner loadext --ibconnection %TESTCONN% --extension YellowQuery --file "%CFEPATH%" --updatedb
pause
