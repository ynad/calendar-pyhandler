@echo off
echo SETUP calendar-pyhandler
echo ------------------------

set /p choice1="Do you want to set up a new python venv (optional but recommended)? (y/n): "

if /i "%choice1%"=="y" (
    echo You chose to use venv. Default directory is '.venv'
    set /p "venv_user=Input a different directory if you want to change it or just press enter: "

    if "%venv_user%"=="" (
        echo No directory specified, using default: .venv
        echo Setting up venv...
        python -m venv ".venv"
        
        echo Activating venv...
        call ".venv\Scripts\activate.bat"
    
    ) else (
        echo Directory user specified: %venv_user%
        echo Setting up venv...
        python -m venv "%venv_user%"
        
        echo Activating venv...
        call "%venv_user%\Scripts\activate.bat"
    )
    pause

) else if /i "%choice1%"=="n" (
    set /p choice2="Do you want to source an existing python venv? (y/n): "

    if /i "%choice2%"=="yes" (
        set /p venv_dir2="Default directory is '.venv'. Input a different directory if you want to change it or just press enter: "

        if "%venv_dir2%"=="" (
            set venv_dir2=.venv
            echo No directory specified, using default: .venv
        )

        if exist "%venv_dir2%\Scripts\activate.bat" (
            echo Activating existing venv...
            call "%venv_dir2%\Scripts\activate"
        ) else (
            echo Cannot locate venv in this path, exiting...
            exit /b 1
        )
    ) else (
        echo NOT using venv
    )
) else (
    echo Invalid input. Please answer 'y' or 'n'
    exit /b 1
)

echo Installing pip requirements...
python -m pip install -r requirements.txt

echo All done!
pause
