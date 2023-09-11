import json
import sys

searchName = sys.argv[1]
filename = 'lecturers-all.json' 

lecturers = []
lecturerNames = []

with open(filename, 'r') as fcc_file:
    data = json.load(fcc_file)
    lecturers.extend(data)

for l in lecturers:
    lecturerNames.append(l['name'])

def searchIn(strings, value):
    res = []
    for s in strings:
        if value in s:
            res.append(s)
    return res

def searchFullName(lecturer):
    matches = searchIn(lecturerNames, lecturer['lastname'])
    if len(matches) == 1:
        return matches[0]
    elif len(matches) > 1:
        secondMatches = [] 
        for l in matches:
            nameParts = l.split(' ')
            if lecturer['n'] in nameParts[1] and lecturer['p'] in nameParts[2]:
                secondMatches.append(l)
        if len(secondMatches) > 0:
            return secondMatches[0]
        else:
            return matches[0]


    return lecturer['lastname'] + ' ' + lecturer['n'] + '. ' + lecturer['p'] + '.'