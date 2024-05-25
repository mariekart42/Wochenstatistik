#!/bin/bash

source_file="/Users/mariemensing/data/Daten_Wochenstatistik.xlsx"
destination_folder="/Users/mariemensing/RiderProjects/Wochenstatistik/Wochenstatistik/document"

if [ -f "$source_file" ]; then
    cp "$source_file" "$destination_folder"
    echo "File copied successfully."
else
    echo "Source file does not exist."
fi