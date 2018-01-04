# DragonBoatLineupSimulator

Description:
Uses paddler data from an Excel spreadsheet and automatically sorts paddlers into simulated boat lineups. 
The results are outputted to a newly generated Excel spreadsheet.

Programs:
The program is written using Python.
Two scripts are available for use - DefaultProgram.py and Program.py.


DefaultProgram.py:
The script automatically takes a Excel spreadsheet named "List.xlsx" as input, sets up 2 boats, each with 20 paddlers and 4 substitutes and a 10 Male to 10 Female gender split. 
The results are outputted to a generated spreadsheet named "DefaultResults.xls".

Programp.py:
This script requires user input. 
The user can select what spreadsheet they want as input (with the .xlsx extension included).
They can select the total number of paddlers being inputted, number of boats to be generated, and the gender ratio for male:female (10:10 or 12:6) for each boat.
The results are automatically saved to a generated spreadsheet named "Results.xls".
