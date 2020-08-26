import re
def MBExtract(string):
    pattern = re.compile(r'.hence.(North|South|East|West|N|S|E|W|north|south|east|west).\d{1,4}.{0,2}\d{0,2}.{0,2}\d{0,2}.{0,2}(North|South|East|West|N|S|E|W|north|south|east|west){0,5}.{0,1}\d{0,4}.{0,1}\d{0,2}.(feet|rods|chains|Feet|ft|Rods)')
    matches = pattern.finditer(string)
    matches1 = ''
    matchlist = []
    for match in matches:
        matches1 += '\n' + match.group()
        matchlist.append(match.group())
    return matches1

#MBExtract(r'Section 22: Beginning 30.00 feet East of the Nortwest corner of the NE/4NE/4; running thence East 105.00 feet; thence South 255.00 feet; thence West 105.00 feet; thence North 255.00 feet to the point of beginning.')