def mbDegreesImport(cDirection1, dDegrees, dMinutes, dSeconds, cDirection2, feet):
    # FXN: Several singular conversions of Degrees/Minutes/Seconds to Decimal Degrees with export
    # INPUTS: Legal description actionable item

    # Format the inputs and outputs
    if (cDirection1 == '') and (dDegrees == '') and (dMinutes == '') and (dSeconds == '') and (cDirection2 == '') and (feet == ''):
        output = [0,0]
    elif (cDirection1 != "N") and (cDirection1 != "n") and (cDirection1 != "S") and (cDirection1 != "S"):
        output = "The Northing value has to be either North or South."
    elif (cDirection2 != "E") and (cDirection2 != "e") and (cDirection2 != "W") and (cDirection2 != "w"):
        output = "The Easting value has to be either East or West."
    elif (cDirection1 == '') and (dDegrees == '') and (dMinutes == '') and (dSeconds == '') and (cDirection2 == '') and (feet == ''):
        output = np.array((0,0))
    else:
        # Convert input to decimal degrees
        dDegrees = float(dDegrees)
        dMinutes = float(dMinutes)
        dSeconds = float(dSeconds)
        feet = float(feet)

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

        output = np.array((decimalDegrees, feet))


    return output
