# Word Batch Replace (C++)

This tool replaces a keyword across all `.doc` and `.docx` files under a folder using Microsoft Word COM automation.
It is intended for Windows and can be built with Dev-C++.

## Requirements
- Windows with Microsoft Word installed.
- Dev-C++ configured to compile C++17.
- Update the `#import` path in `src/main.cpp` to match your Office installation (MSWORD.OLB).

Common MSWORD.OLB paths:
- `C:\\Program Files\\Microsoft Office\\root\\Office16\\MSWORD.OLB`
- `C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\MSWORD.OLB`

## Build (Dev-C++)
1. Open `src/main.cpp` in Dev-C++.
2. Set the compiler standard to C++17.
3. Build the project to produce `WordBatchReplace.exe`.

## Usage
```
WordBatchReplace.exe <folder> <find> <replace>
```

Example:
```
WordBatchReplace.exe "D:\\Docs" "旧关键词" "新关键词"
```

## Notes
- The program will recurse into subfolders.
- Files are only saved if replacements are found.
- Word UI stays hidden while running.
