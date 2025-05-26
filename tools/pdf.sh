#!/bin/bash
set -euo pipefail

ROOT="$(dirname "$0")"
cd "$ROOT/../books"
for file in ./*.docx; do
    if [[ "$(basename "$file")" != ~* ]]; then
        echo "$file"
    fi
    
    pdf_file="${file%.docx}.pdf"
    if [[ ! -f "$pdf_file" ]]; then
        echo "Converting $file to $pdf_file"
        "/Applications/LibreOffice.app/Contents/MacOS/soffice" --headless --convert-to pdf "$file" --outdir .
    else
        echo "$pdf_file already exists, skipping."
    fi 
    
    
    epub_file="${file%.docx}.epub"
    if [[ ! -f "$epub_file" ]]; then
        echo "Converting $file to $epub_file"
        "/Applications/LibreOffice.app/Contents/MacOS/soffice" --headless --convert-to epub "$file" --outdir .
    else
        echo "$epub_file already exists, skipping."
    fi 
done

# libreoffice --headless --convert-to pdf /Users/ha/word/input.docx --outdir /Users/ha/word/output
# libreoffice --headless --convert-to epub /Users/ha/word/input.docx --outdir /Users/ha/word/output