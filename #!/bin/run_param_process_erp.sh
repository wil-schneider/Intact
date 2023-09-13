#!/bin/bash

# Check if the number of arguments provided is correct
if [ "$#" -ne 3 ]; then
    echo "Usage: ./run_python.sh [DIRECTORY_PATH] [FILENAME] [OUTPUT_NAME]"
    exit 1
fi

DIRECTORY_PATH="$1"
FILENAME="$2"
OUTPUT_NAME="$3"

# Run the Python script with the provided arguments
python parameterized.py -b "$DIRECTORY_PATH" -f "$FILENAME" -o "$OUTPUT_NAME"
