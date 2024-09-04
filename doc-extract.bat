@echo off
setlocal enabledelayedexpansion
echo  -- doc-extract v1.0.0 -- 
echo     a project by james.m
echo 
rem List all .docx files in the current directory
echo List of .docx files:
set count=0
for %%f in (*.docx) do (
    set /a count+=1
    echo !count!. %%f
    set "file!count!=%%f"
)

rem Check if there are no .docx files
if %count%==0 (
    echo No .docx files found in the current directory.
    pause
    exit /b
)

rem Add option for "All of the above"
set /a count+=1
echo %count%. All of the above

rem Prompt the user to select a file by number or choose "All"
set /p choice="Select a file by entering the corresponding number (or %count% for all): "

rem Validate user input
if %choice% lss 1 (
    echo Invalid choice.
    pause
    exit /b
)

if %choice% gtr %count% (
    echo Invalid choice.
    pause
    exit /b
)

if %choice%==%count% (
    rem Process all .docx files
    for %%f in (*.docx) do (
        call :process_file "%%f"
    )
) else (
    rem Get the selected file
    call :process_file "!file%choice%!"
)

pause
exit /b

:process_file 
set "selected_file=%~1"
set "file_base_name=%selected_file:~0,-5%"
echo Processing %selected_file%...

rem Convert the selected .docx file to a .zip file
ren "%selected_file%" "%file_base_name%.zip"
echo Converted %selected_file% to zip format.

rem Extract the "media" folder from the zip file
echo Extracting "media" folder...
set "zip_file=%file_base_name%.zip"
set "output_dir=%file_base_name%_media"

rem Use PowerShell to extract the "media" folder
powershell -command "Expand-Archive -Path '%cd%\%zip_file%' -DestinationPath '%cd%\%output_dir%' -Force"
if not exist "%output_dir%\word\media" (
    echo "media" folder not found in %zip_file%.
    exit /b
)
move "%output_dir%\word\media\*" "%output_dir%"
rd /s /q "%output_dir%\word"
rd /s /q "%output_dir%\_rels"
rd /s /q "%output_dir%\customXml"
rd /s /q "%output_dir%\docProps"
del "%output_dir%\[Content_Types].xml"

echo "media" folder extracted to %output_dir%.

rem Rename image files in the output directory
cd "%output_dir%"
set count=0
for %%i in (*.jpg *.png *.jpeg *.gif) do (
    set /a count+=1
    ren "%%i" "%file_base_name%_image!count!%%~xi"
)
cd ..

rem Convert the .zip file back to .docx
ren "%zip_file%" "%file_base_name%.docx"
echo Converted %zip_file% back to docx format.

exit /b
