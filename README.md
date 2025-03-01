# Student Course Placement and Scheduling System

## Overview
This project is a Python-based course placement and scheduling system designed specifically for Hamilton College Registrar's Office. The script processes student course preferences, placements, and constraints to optimize class assignments using Gurobi optimization.

## Features
- **CSV Data Processing**: Reads and processes class information from `classes.csv` and student preferences from `priorities.csv`.
- **Class and Student Data Structuring**: Organizes class schedules, department classifications, and student preferences.
- **Gurobi Optimization Model**: Implements constraints such as course limits, overlapping schedules, and priority-based selections.
- **Automatic Placement Handling**: Adjusts student placements based on predefined rules and AP exam scores.
- **Result Export**: Generates output files including:
  - `results.txt`: Statistics on the optimization run
  - `result.xlsx`: Structured output of student schedules for analysis
  - Additional reports on class availability and student preferences

## Prerequisites
- Python 3.11
- Required libraries:
  - `pandas`
  - `gurobipy`
  - `csv`
- Gurobi Optimizer installed and licensed

## Installation
1. Install dependencies using:
   ```sh
   pip install pandas gurobipy
   ```
2. Ensure Gurobi is installed and licensed on your system.

## Usage
1. Prepare the input CSV files:
   - `classes.csv` (course information)
   - `priorities.csv` (student preferences and placements)
Please refer to the csv templates in this repository.
2. Run the script:
   ```sh
   python main.py
   ```
3. View the generated output files for results.

## Configuration
- Modify constants in `main.py` to adjust weightings, constraints, or department-specific rules.
- Update the `PLACEMENTS` dictionary to include new placement rules.

## Output Files
- **Schedules and Preferences**:
  - `result.xlsx`: Structured output of student schedules
  - `students.txt`: Recorded student preferences
- **Statistical Analysis**:
  - `results.txt`: Summary statistics
- **Course Information**:
  - `classes.txt`: Class meeting information
  - `multisection.txt`: Multi-section course details
  - `crosslisted.txt`: Cross-listed course details
  - `departments.txt`: Department-wise course distribution
  - `unique_placements.txt`: Identified unique placement cases

## Notes
- Ensure input CSV files are correctly formatted to avoid errors.
- Gurobi license issues should be resolved before execution.
- Debugging logs can be enhanced by adding print statements or logging.

## Authors
- **Project Maintainer:** Grisha Hatavets
  
## License
This project is licensed under MIT License.


