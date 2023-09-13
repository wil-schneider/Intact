#!/bin/bash

# Check if the file is provided
if [ -z "$1" ]; then
    echo "Usage: $0 <CSV file>" >&2
    exit 1
fi

# Read the file line by line
count=0
while IFS= read -r line
do
    # Only process the first two lines
    ((count++))
    if [ $count -gt 2 ]; then
        break
    fi

    # Get header line to find the index of "Username"
    if [ $count -eq 1 ]; then
        IFS=','
        header=($line)
        for i in "${!header[@]}"; do
            if [ "${header[$i]}" == "Username" ]; then
                username_idx=$i
            fi
        done
        if [ -z "$username_idx" ]; then
            echo "Username column not found in the CSV file." >&2
            exit 1
        fi
    fi

    # Get the first data line and echo the username value to stderr
    if [ $count -eq 2 ]; then
        IFS=','
        data=($line)
        echo "Username is: ${data[$username_idx]}" >&2
    fi
done < "$1"
