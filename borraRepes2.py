#!/usr/bin/env python3.7

import sys

def sameRow(concerningRow, rowToCompare, indices):
    if not indices:
        for i, v in enumerate(concerningRow):
            if v != rowToCompare[i]:
                return False
    else:
        for i in indices:
            if concerningRow[i] != rowToCompare[i]:
                return False

        return True

with open(sys.argv[1], "r") as fIn, open(sys.argv[2], "w") as fOut:
    firstRowColumnsNames = fIn.readline()
    fOut.write(firstRowColumnsNames)

    #for i, v in enumerate(firstRowColumnsNames.split(",")):
    #    print("{}: {}".format(i + 1, v))

    #referenceColumns = input("Select reference columns separated by ','.\n\
#If references are not set it will be all the columns: ").split(",")
    referenceColumns = [14, 17, 19, 20]
    #referenceColumns = [15, 18, 20, 21]

    concerningRow = fIn.readline()
    fOut.write(concerningRow)

    i = 2
    concerningLine = 2

    for line in fIn:
        if sameRow(concerningRow.split(","), line.split(","), referenceColumns) is False:
            concerningLine = i
            concerningRow = line
            fOut.write(line)
        else:
            print("{} and {} lines are identical.".format(concerningLine, i))

        i += 1
