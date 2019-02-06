## Spreadsheet Finder

Quick'n dirty find spreadsheet files containing specific values matching your regular expression.

Example: 
```
cargo run --release -- -r "Umlaute" "/Temp/Korrekturen"
```

Finds all ods files containing "Umlaute" as cell value.
You can specify the matcher for filenames also (default value used here):
```
cargo run --release -- --fileregex ".*\.ods" "-r "Umlaute" "/Temp/Korrekturen"
```
