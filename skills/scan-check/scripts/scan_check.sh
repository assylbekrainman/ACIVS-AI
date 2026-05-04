#!/bin/bash

# Simple wrapper for the verify_document.py script

if [ "$#" -ne 1 ]; then
    echo "Usage: $0 <path_to_document>"
    exit 1
fi

python3 /home/ubuntu/skills/scan-check/scripts/verify_document.py "$1"
