#!/bin/bash

source_file="/Users/mariemensing/data/Daten_Wochenstatistik.xlsx"
destination_folder="/Users/mariemensing/RiderProjects/Wochenstatistik/Wochenstatistik/document"

# Check if the destination folder exists, if not, create it
if [ ! -d "$destination_folder" ]; then
    mkdir -p "$destination_folder"
    echo "Destination folder created."
fi

# Check if the source file exists, if so, copy it
if [ -f "$source_file" ]; then
    cp "$source_file" "$destination_folder"
    echo "File copied successfully."
else
    echo "Source file does not exist."
fi