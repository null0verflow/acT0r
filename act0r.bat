

@echo off

:STP
Set _os_bitness=64
IF %PROCESSOR_ARCHITECTURE% == x86 (
  IF NOT DEFINED PROCESSOR_ARCHITEW6432 Set _os_bitness=32
  )
IF NOT %_os_bitness% ==64 goto FAIL
cd ..
cd ..
cd ..
cd ..   

C:
    IF EXIST C:\"Program files (x86)"\"Microsoft Office"\Office16\ GOTO CON
    CD \"Program files (x86)"
   IF NOT EXIST C:\"Program files (x86)"\"Microsoft Office"\Office16\NUL GOTO MIS
   CD \"Program files (x86)"

:CON
echo copyright @bobdinh139 
echo thank you for using this program (this is just a beta), pls feedback! 
echo This could fail multiple times, try again (for office 2016, 64bit devices only) (remember to close office and turn of wifi) 
echo Loading.... 
ping n- 2 127.0.0.1>nul
echo Press any key to continue:
set/p input=
cd ..
cd ..
cd ..
cd ..
cd ..
cd Program files (x86)
cd Microsoft Office
cd Office16
echo Type in your activation key (New key or used key is okay)(it should looks like this: XXXXX-XXXXX-XXXXX-XXXXX-XXXX.) 
echo if you type wrong, this will not allow you to type again (restart the program)-missing function!
echo To use provided key (which is used before), type "p":
set/p "type=>"

if %type% ==p goto PO
if NOT %type% ==p goto NP

:NP
cscript OSPP.VBS /inpkey: %type% 
cscript OSPP.VBS /sethst:kms.digiboy.ir
cscript OSPP.VBS /act
cscript OSPP.VBS /dstatus

:PO
cscript OSPP.VBS /inpkey: XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 
cscript OSPP.VBS /sethst:kms.digiboy.ir
cscript OSPP.VBS /act
cscript OSPP.VBS /dstatus
echo if the program fail, try again or try to run as admin, buy microsoft office or use free trail here (press b to go to, c to cancel, a to try again as admin)
echo if you type wrong, this will not allow you to type again (restart the program)-missing function!
set/p "web=>"

if %web% ==b goto WEB
if %web% ==c goto DONE
if %web% ==a goto RAD

:WEB
start "" https://www.mychoicesoftware.com/pages/office-2016-home-and-student?gclid=CjwKCAiAhMLSBRBJEiwAlFrsTn6-Wh0q32-kPAkD06UJCKGcOeN6ND5fpQlHlgmkS15Vxk73hLR1dRoCsE4QAvD_BwE

:FAIL
echo x=msgbox("You are using 32-bit device!" , 0+16 , "Failed!") >message.vbs
start message.vbs

color 2
timeout 0 > nul
cd Desktop
del message.vbs
pause
exit

:MIS
echo missing folders!
echo x=msgbox("Folder missing!" , 0+64 , "Failed!") >message.vbs
start message.vbs

color 2
timeout 0 > nul
cd Desktop
del message.vbs
pause
exit

:DONE
pause

:RAD
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"

    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    pushd "%CD%"
    CD /D "%~dp0"

 goto STP


