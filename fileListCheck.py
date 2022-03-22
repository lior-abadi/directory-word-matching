import Levenshtein as lv
import os
import pandas as pd

# This program first finds which fields of compareDirs are located in the general file of allDirs.
# Once each match is detected, creates a new outputDir file with the mismatched data.

allDirs = pd.read_excel("inputDirWithAllEntries.xlsx", engine="openpyxl", dtype=object, header = 0)
compareDirs = pd.read_excel("comparisonData.xlsx", engine="openpyxl", dtype=object, header = 0)
outputDir = allDirs

distTolerance = 5 # Amount of characters that differ.
jaroTolerance = 0.92 # The closer to 1, the stronger the match.

# Aux function if the detection of a specific char is desired. 
# Where string is the chain and char the character to identify.
def charPosition(string, char):
    position = []
    for n in range(len(string)):
        if string[n] == char:
            position.append(n)
    return position

matches  = 0
print("Looping over ", len(allDirs), " base data and comparing with ", len(compareDirs), " data.")
for row in range(len(allDirs)): 
    allDir = allDirs["SET_A_COLUMN"][row].strip() # Set the column to read
    print("Evaluating registry: ", allDir)
    for compare in range(len(compareDirs)):
            compareDir = compareDirs["SET_A_COLUMN2"][compare].strip()            
            jaroDiff = lv.jaro(compareDir, allDir)
            distance = lv.distance(compareDir, allDir)
            
            if (distance < distTolerance and jaroDiff>jaroTolerance):
                matches += 1
                outputDir["SET_OUTPUT_COLUMN"][row] = True

                
outputDir.to_excel("output.xlsx")
print("Matches: ", matches)


