# OCROK
Program to help ROK R4 to analyse statistics of their members.

# How it works ?
## Installation
- step 1 : download and install : https://github.com/UB-Mannheim/tesseract/wiki
- step 2 : download the last release of OCROK

## Use it :
- step 0 : your game must be in english
- step 1 : screenshot all detailed information of your members (sorry, this is still a manual task), like this one :
![alt text](https://raw.githubusercontent.com/lysm-rok/OCROK/main/pictures/screen1.jpg)
- step 2 : put all the files in the same fodler
- run the program and follow the instructions
- it will (should) create a csv file, with all information of the screens

As I use the https://github.com/MScholtes/PS2EXE tool to create the executable file (which is more user friendly), it might raise an wrong antivirus alert.. 
I commited the script itself, so you can choose to execute it directly from a powershell console.

## Why?
Beceause I spent too much time scrolling in the game and handtyping in excel to attributes kvk rewards to best players. 

## What is in the box?
Technically speaking : 
- It uses tesseract software
- There is a powershell wrapper / parser packaged in a .exe file
- there is a tiny .database.rok file created at the first run. It is used to improved the analysis for the next runs.

/!\ please note that I am not a programmer, there are hundreds of better ways to program it, but it seems to work for me. 
