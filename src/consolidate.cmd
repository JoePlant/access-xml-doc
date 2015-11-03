rem @echo off
set database=database.xml
set dir=.\xml

if EXIST Output goto Output_exists
mkdir Output
:Output_exists

if EXIST Working goto Working_exists
mkdir Working
:Working_exists

set nxslt=..\lib\nxslt\nxslt.exe

@echo === Apply Templates ===
%nxslt% %dir%\%database% Stylesheets\consolidate.xslt -o Working\database.xml 2> working\errors.txt 
