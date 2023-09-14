#!/bin/bash

# Base filepath
base_filepath="/var/data/fin_ifrs/erp/inbound/processing/gl_bu/"

# Check if a filename is provided
if [ -z "$1" ]; then
  echo "Usage: $0 <filename>" >&2
  exit 1
fi

# Concatenate input filename with "_username.txt" and the base filepath
file_to_read="${base_filepath}${1}_username.txt"

# Check if file exists
if [ ! -f "$file_to_read" ]; then
  echo "Error: File '$file_to_read' not found" >&2
  exit 1
fi

# Read and pass the file content to stderr
cat "$file_to_read" >&2
