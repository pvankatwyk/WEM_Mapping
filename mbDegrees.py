def mbDegrees(cDirection1, dDegrees, dMinutes, dSeconds, cDirection2, feet, unit):
    # FXN: Convert input to decimal degrees
    # INPUTS: Legal description actionable item

    # Make sure inputs are floats
    dDegrees = float(dDegrees)
    dMinutes = float(dMinutes)
    dSeconds = float(dSeconds)
    feet = float(feet)

    # Convert degrees
    degrees = dDegrees + (dMinutes / 60) + (dSeconds / 3600)
    decimalDegrees = 1
    # Configure the decimal degrees according to directionality
    if (cDirection1 == 'N' or cDirection1 == 'n') and (cDirection2 == 'E' or cDirection2 == 'e'):
        decimalDegrees = 0 + degrees
    elif (cDirection1 == 'N' or cDirection1 == 'n') and (cDirection2 == 'W' or cDirection2 == 'w'):
        decimalDegrees = 360 - degrees
    elif (cDirection1 == 'S' or cDirection1 == 's') and (cDirection2 == 'E' or cDirection2 == 'e'):
        decimalDegrees = 180 - degrees
    elif (cDirection1 == 'S' or cDirection1 == 's') and (cDirection2 == 'W' or cDirection2 == 'w'):
        decimalDegrees = 180 + degrees

    # Convert units into feet
    if unit == 'feet':
        feet = feet
    elif unit == 'rods':
        feet = feet * 16.5
    elif unit == 'chains':
        feet = feet * 66

    output = str(decimalDegrees) + ' degrees ' + str(feet) + ' feet'

    # Format some error messages (to catch wrong inputs)
    if (cDirection1 != "N") and (cDirection1 != "n") and (cDirection1 != "S") and (cDirection1 != "S"):
        output = "The Northing value has to be either 'N' or 'S'."
    if (cDirection2 != "E") and (cDirection2 != "e") and (cDirection2 != "W") and (cDirection2 != "w"):
        output = "The Easting value has to be either 'E' or 'W'."

    return output