def MBUltimate(MBstring):
    import re
    import pandas as pd
    # FXN: Takes a legal description string, separates it into actionable steps, and exports the results into CSV to be uploaded into QGIS
    # INPUT: Legal Description (string)

    # Identify pattern that will find actionable steps
    pattern = re.compile(r'.hence.(North|South|East|West|N|S|E|W|north|south|east|west).\d{1,4}.{0,2}\d{0,2}.{0,2}\d{0,2}.{0,2}(North|South|East|West|N|S|E|W|north|south|east|west){0,5}.{0,1}\d{0,4}.{0,1}\d{0,2}.(feet|rods|chains|Feet|ft|Rods)')
    matches = pattern.finditer(MBstring)
    matches1 = ''
    matchlist = []
    # Compile the matches into a list
    for match in matches:
        matches1 += '\n' + match.group()
        matchlist.append(match.group())

    match_split = []

    # Split the matches on white space (\s), or degrees minutes seconds notation
    for item in matchlist:
        match_split.append((re.split(r"[\s°’']", item)))

    # For actions which only have 2 components (such as West 100 feet), created empty strings to show there are no degrees minutes seconds
    for i in range(len(match_split)):
        if len(match_split[i]) < 5 and (match_split[i][1] == 'North' or match_split[i][1] =='N' or match_split[i][1] =='n' or match_split[i][1] =='north' or match_split[i][1] =='South' or match_split[i][1] =='s' or match_split[i][1] =='south' or match_split[i][1] =='S'):
            match_split[i].insert(2,'')
            match_split[i].insert(3, '')
            match_split[i].insert(4, '')
            match_split[i].insert(5, '')
        elif len(match_split[i]) < 5 and (match_split[i][1] == 'East' or match_split[i][1] =='E' or match_split[i][1] =='e' or match_split[i][1] =='east' or match_split[i][1] =='West' or match_split[i][1] =='W' or match_split[i][1] =='west' or match_split[i][1] =='w'):
            match_split[i].insert(1, '')
            match_split[i].insert(2, '')
            match_split[i].insert(3, '')
            match_split[i].insert(4, '')

    # Make the list of lists a data frame (makes it easier to work with)
    df_matches = pd.DataFrame(match_split)
    # Name the columns
    df_matches.columns = ['thence', 'northing', 'degree', 'minute', 'second', 'easting', 'distance', 'unit']
    # Drop the column with all of the 'thence' in it
    df_final = df_matches.drop(columns = ['thence'])

    # Convert all of the numbers into numeric values in order to do calculations
    df_final['degree'] = pd.to_numeric(df_final['degree'], errors = 'coerce')
    df_final['minute'] = pd.to_numeric(df_final['minute'], errors = 'coerce')
    df_final['second'] = pd.to_numeric(df_final['second'], errors = 'coerce')
    df_final['distance'] = pd.to_numeric(df_final['distance'], errors = 'coerce')

    # Fill in 0 in all of the spaces with "NA"
    df_final.fillna(0, inplace = True)

    # Conversion calculation into new 'degrees' column (not working with directionality)
    df_final['degrees'] = df_final['degree'] + (df_final['minute'] / 60) + (df_final['second'] / 3600)

    # Get new columns ready (pre-allocate them and make sure they're floats)
    df_final['decimalDegrees'] = 0
    df_final['decimalDegrees'] = pd.to_numeric(df_final['decimalDegrees'], errors = 'coerce')
    df_final['decimalDegrees'] = df_final['decimalDegrees'].astype(float)
    df_final['distanceConverted'] = 0
    df_final['distanceConverted'] = df_final['distanceConverted'].astype(float)

    # Encorporate directionality into conversion above
    #   - basically, if North, add 0 / South, add 180 to the calculation. Then, if East then add degrees (clockwise), west subtract (counterclockwise)
    for i in range(len(df_final)):
        if ((df_final['northing'][i] == "N") or (df_final['northing'][i] == 'n') or (df_final['northing'][i] == 'North') or (df_final['northing'][i] == 'north')) and ((df_final['easting'][i] == 'East') or (df_final['easting'][i] == 'east') or (df_final['easting'][i] == 'E') or (df_final['easting'][i] == 'e')):
            df_final['decimalDegrees'][i] = 0 + df_final['degrees'][i]
        elif ((df_final['northing'][i] == 'N') or (df_final['northing'][i] == 'n') or (df_final['northing'][i] == 'North') or (df_final['northing'][i] == 'north')) and ((df_final['easting'][i] == 'W') or (df_final['easting'][i] == 'w') or (df_final['easting'][i] == 'West') or (df_final['easting'][i] == 'west')):
            df_final['decimalDegrees'][i] = 360 - df_final['degrees'][i]
        elif ((df_final['northing'][i] == 'S') or (df_final['northing'][i] == 's') or (df_final['northing'][i] == 'South') or (df_final['northing'][i] == 'south')) and ((df_final['easting'][i] == 'E') or (df_final['easting'][i] == 'e') or (df_final['easting'][i] == 'East') or (df_final['easting'][i] == 'east')):
            df_final['decimalDegrees'][i] = 180 - df_final['degrees'][i]
        elif ((df_final['northing'][i] == 'S') or (df_final['northing'][i] == 's') or (df_final['northing'][i] == 'South') or (df_final['northing'][i] == 'south')) and ((df_final['easting'][i] == 'W') or (df_final['easting'][i] == 'w') or (df_final['easting'][i] == 'West') or (df_final['easting'][i] == 'west')):
            df_final['decimalDegrees'][i] = 180 + df_final['degrees'][i]
        elif ((df_final['northing'][i] == 'S') or (df_final['northing'][i] == 's') or (df_final['northing'][i] == 'South') or (df_final['northing'][i] == 'south')) and df_final['easting'][i] == '':
            df_final['decimalDegrees'][i] = 180
        elif (df_final['northing'][i] == '') and ((df_final['easting'][i] == 'W') or (df_final['easting'][i] == 'w') or (df_final['easting'][i] == 'West') or (df_final['easting'][i] == 'west')):
            df_final['decimalDegrees'][i] = 270
        elif (df_final['northing'][i] == '') and ((df_final['easting'][i] == 'E') or (df_final['easting'][i] == 'e') or (df_final['easting'][i] == 'East') or (df_final['easting'][i] == 'east')):
            df_final['decimalDegrees'][i] = 90

    # Convert units into feet
    for i in range(len(df_final)):
        if df_final['unit'][i] in ('feet', 'ft', 'Feet', 'ft.', 'Ft'):
            df_final['distanceConverted'][i] = df_final['distance'][i]
        elif df_final['unit'][i] in ('rods', 'Rods'):
            df_final['distanceConverted'][i] = df_final['distance'][i] * 16.5
        elif df_final['unit'][i] == 'chains':
            df_final['distanceConverted'][i] = df_final['distance'][i] * 66

    # Copy to new dataframe the columns we need
    df_decimal = df_final[['decimalDegrees','distanceConverted']].copy()

    # Export as a CSV in a format that the M&B mapper on QGIS will recognize it
    pd.DataFrame.to_csv(self = df_decimal, path_or_buf=r'C:\Users\Accounting\Downloads\MBUltimate.csv',
                                                encoding='utf-8', index=False, sep=';', header = False)
    return matches1