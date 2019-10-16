cls
:: -----------------------------------------------------------------
:: April 22nd 2002
:: Visual FoxPro Toolkit for .NET Installation
:: Dynamic Help:    Check C,D,E and F folders for a specific folder and copies the XML file
:: Update Registry: Updates the registry so the VFPToolkit appears in the "Add References" dialog
:: -----------------------------------------------------------------

::-- Clear the screen
cls
@ECHO OFF

::-- Check the C drive
:CHECK_C
IF NOT EXIST "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\1033" GOTO CHECK_D
C: 
COPY VFPToolkitNET.xml "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\1033\VFPToolkitNET.xml"
ECHO.
GOTO END 

::-- Check the D drive
:CHECK_D
IF NOT EXIST "D:\program files\Microsoft Visual Studio .NET\Common7\IDE\HTML\XMLLinks\1033" GOTO CHECK_E
D: 
COPY VFPToolkitNET.xml "D:\program files\Microsoft Visual Studio .NET\Common7\IDE\HTML\XMLLinks\1033\VFPToolkitNET.xml"
ECHO.
GOTO END 

::-- Check the E drive
:CHECK_E
IF NOT EXIST E:\program files\Microsoft Visual Studio .NET\Common7\IDE\HTML\XMLLinks\1033 GOTO CHECK_F
E: 
COPY VFPToolkitNET.xml "E:\program files\Microsoft Visual Studio .NET\Common7\IDE\HTML\XMLLinks\1033\VFPToolkitNET.xml"
ECHO.
GOTO END 

::-- Check the F drive
:CHECK_F
IF NOT EXIST F:\program files\Microsoft Visual Studio .NET\Common7\IDE\HTML\XMLLinks\1033 GOTO END
F: 
COPY VFPToolkitNET.xml "F:\program files\Microsoft Visual Studio .NET\Common7\IDE\HTML\XMLLinks\1033\VFPToolkitNET.xml"
ECHO.
GOTO END 

::-- Done
:END

::-- Now call the reg file to update the registry
VFPToolkitNET.reg

::-- Open the readme file
notepad readme.txt 

@ECHO ON
EXIT


::-- Temp code
::REM [For the Final Version the above remark will be removed.]