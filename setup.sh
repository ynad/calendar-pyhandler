#!/bin/bash

printf "SETUP calendar-pyhandler\n-------------------------\n\n"

read -p "Do you want to set up a new python venv (optional but recommended)? (Yes/No): " choice1

case "$choice1" in
    yes|y|Y|Yes|YES)
        printf "You chose to use venv. Default directory is '.venv'\n"
        read -p "Input a different directory if you want to change it or just press enter: " venv_dir

        # Check if the input is empty
        if [[ -z "$venv_dir" ]]; then
            printf "No directory specified, using default: .venv\n\n"
            venv_dir=".venv"
        else
            printf "Directory set to: $venv_dir\n\n"
        fi

        # activate python venv
        printf "Setting up venv...\n"
        python3 -m venv "$venv_dir"

        # source it
        source "$venv_dir/bin/activate"

        ;;

    no|n|N|No|NO)
        read -p "Do you want to source an existing python venv? (Yes/No): " choice2
        case "$choice2" in
            yes|y|Y|Yes|YES)
                read -p "Default directory is '.venv'. Input a different directory if you want to change it or just press enter: " venv_dir2

                # Check if the input is empty
                if [[ -z "$venv_dir2" ]]; then
                    printf "No directory specified, using default: .venv\n\n"
                    venv_dir2=".venv"
                else
                    printf "Directory set to: $venv_dir2\n\n"
                fi

                # source it
                printf "Source '$venv_dir2' dir\n"
                source "$venv_dir2/bin/activate"

                if [ $? -ne 0 ]; then
                    echo Cannot locate venv in this path, exiting...
                    exit 1
                fi

                ;;

            no|n|N|No|NO)
                printf "NOT using venv\n\n"
                ;;

            *)
                printf "No choice specified, NOT using venv\n\n"
                ;;
        esac
        ;;
    *)
        echo "Invalid input. Please answer 'yes/y/no/n'"
        exit 1
        ;;
esac

# install pip requirements
printf "\nInstalling pip requirements...\n\n"
python3 -m pip install -r requirements.txt

printf "\nAll done!\n"
