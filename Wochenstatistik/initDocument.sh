#!/bin/bash

excel_file="/Users/mariemensing/data/Daten Wochenstatistik.xlsx"
user_file="/Users/mariemensing/data/Nutzer Liste.txt"
document_folder="/Users/mariemensing/RiderProjects/Wochenstatistik/Wochenstatistik/document"

if [ ! -d "$document_folder" ]; then
    mkdir -p "$document_folder"
    echo "Destination folder created."
fi

if [ -f "$excel_file" ]; then
    cp "$excel_file" "$document_folder"
    echo "File copied successfully."
else
    echo "Source file does not exist."
fi

if [ -f "$user_file" ]; then
    cp "$user_file" "$document_folder"
    echo "File copied successfully."
else
    echo "Source file does not exist."
fi