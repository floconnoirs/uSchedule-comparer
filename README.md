# uSchedule-comparer

A tool that visually compares JSON format schedules exported from [uSchedule.me](https://uSchedule.me) and converts them into an Excel sheet for easy visualization.

## Features
- Compares schedules visually using colors.
- Outputs results in an Excel file with class times highlighted.
- Shows free time (green boxes) and busy class times (red boxes).

## Requirements
- **Excel**
- **Python Interpreter (3.6+)**
- **Git Bash** (or any terminal/command prompt)

## Installation
- Install the `openpyxl` library by running the following command:
   ```bash
   pip install openpyxl

## Usage
- Create schedules using [uSchedule.me](https://uSchedule,me).
- Export the schedules in JSON format.
- Place the exported JSON files into the uSch_JSON folder.
- Move to the folder where the project is located by running:
   ```bash
   cd path/to/uSchedule-comparer
- Execute the Pyhton script by running:
   ```bash
   python uSch_comp.py

## Output
- The results are saved into an Excel file named `schedule.xlsx`
- Each day of the week is seperated into different sheets
- Green boxes represent free time
- Red boxes represent class time
