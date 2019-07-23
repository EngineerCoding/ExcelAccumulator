# Excel Accumulator

A small project to sum values in a single cell across all worksheets in a single excel sheet.

# Running

Before running, please install the dependencies:
```
pip install -r requirements.txt
```

Note that this also installs the PyQt dependency!

This program can be run either as terminal application, by directly calling the accumulator script. If a GUI is desired, run the gui script.

# Windows executable

To build a windows executable for plebs, please run the following:

```
pip install pyinstaller
pyinstaller gui.spec -F
```
