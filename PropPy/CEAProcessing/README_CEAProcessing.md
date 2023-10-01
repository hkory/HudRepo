Hello

This README will cover items in the CEAProcessing Folder

*Link to a video covering AutoCEA.py Usage*

AutoCEA_Input.xlsx
    - This excel sheet is where you will enter all information required to get something useful out of NASA CEA.
    - Further notes on this sheet's use are included within.

AutoCEA_Output.xlsx
    - This excel sheet is where the AutoCEA.py code will spit out the cleaned and formatted CEA Output.
    - From this information, you should be able to have a decent idea of what a good propellant Mass Ratio should be, and the relevant corresponding combustion properties.

AutoCEA.py
    - This is the main script that will take your input file, run it through CEA, do some cleaning, and spit it all out on AutoCEA_Output.xlsx.
    - All you should have to do is run the script. The only thing I could see you maybe needing to change is file paths.
Watch this video to view the functionality of this script - https://youtu.be/RKBDQbuojgI

CEATabulationCleaner.py
    - A function that will take the horrid CEA Tabulation output and make it a useable CSV. In retrospect, I probably should have made the Chamber, Throat, and Exit data all seperate columns. 

CleanCEA_Tabulation.txt
    - Where the cleaned CEA Tabulation data is stored.

RawCleanCEA_Output.txt
    - Where the raw CEA Output data is stored.

RawCEA_Tabulation.txt
    - Where the raw CEA Tabulation data is stored.