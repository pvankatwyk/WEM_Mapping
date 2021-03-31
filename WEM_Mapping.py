# Import Modules
import PySimpleGUI as sg
import pandas as pd
import numpy as np
import os
import re
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementNotInteractableException
from selenium.webdriver import ActionChains
from datetime import datetime
import shutil
import glob
from playsound import playsound
import threading
from MBExtract import MBExtract
from updateLabels import updateLabels
from MBUltimate import MBUltimate
from mbDegreesImport import mbDegreesImport
from mbDegrees import mbDegrees
from iLandmanTract import iLandmanTract
from iLandmanMapTract import iLandmanTractMap
from productionUpdate import productionUpdate
from soundfxn import sound
os.environ['PYTHONHOME'] = r'C:\Users\Accounting\Anaconda3'
os.environ['PYTHONPATH'] = r'C:\Users\Accounting\Anaconda3\lib\site-packages'
os.environ['R_HOME'] = 'C:/Program Files/R/R-4.0.2'
os.environ['R_USER'] = r'C:\Users\Accounting\Anaconda3\lib\site-packages\rpy2'
import rpy2.robjects as ro
# TODO: Add Comments!!!
# TODO: Update the user manual

# APP THEME
sg.theme('Dark2')

# Establish the Main Window Layout
layout_Main = [[sg.Image(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\Maps\Code\WEM (2).png')],
               [sg.Text('What would you like to view?')],
               [sg.Button('Programs'), sg.Button('Files')]]

# Activate the layout and make it a window, which is assigned to win_Main (to be activated later)
win_Main = sg.Window("WEM Mapping Application", element_justification = 'c').Layout(layout_Main)

# Start with windows as False (or not active) and turn them to True (turn them on) when we open them
winActive_Programs = False
winActive_Files = False
winActive_Labels = False
winActive_MB = False
winActive_Layer = False
winActive_MBImport = False
winActive_MBExtract = False
winActive_MBUltimate = False
winActive_AddTract = False
winActive_Prod = False
threadMessage = "Program Started"

# Begin the main window. The way GUI's work, the application runs on while loops. For example:
#       while the application is activated, run it. Once the loop stops, stop the application.
# Note: I will comment heavily the first couple windows (and unique statements later on), but you'll get the idea
while True:
    # Tells the window to read the ev_Main Event (such as a click), and vals_Main values (user input)
    ev_Main, vals_Main = win_Main.Read()
    # If the main event is to Exit, close the while loop, thus closing the program
    if ev_Main is None or ev_Main == 'Exit':
        break
    # Programs Window
    # If the Programs window is not already active and you click (event) Programs....
    if not winActive_Programs and ev_Main == 'Programs':
        # Open the Programs window and hide the main window
        winActive_Programs = True
        win_Main.Hide()
        # Like before, configure the layout for this window and activate it (win_Programs)
        layout_Programs = [[sg.Text('Which Program?')],
                           [sg.Text('Map Updates')],
                           [sg.Button('QGIS Layer Updates')],
                           [sg.Button('QGIS Label Updates')],
                           [sg.Text('Metes and Bounds Calculators')],
                           [sg.Button('Decimal Degrees M&B Calculator')],
                           [sg.Button('M&B String Extractor')],
                           [sg.Button('QGIS M&B Import/Calculator')],
                           [sg.Button('The Ultimate M&B Program')],
                           [sg.Text('iLandman')],
                           [sg.Button('Add Tract')],
                           [sg.Text('Production')],
                           [sg.Button('Update Production Spreadsheet')]]
        win_Programs = sg.Window('Programs', element_justification = 'l').Layout(layout_Programs)

        while True:
            # while in this program, read the event and values, and if the event is close, exit and unhide the last window.
            ev_Programs, vals_Programs = win_Programs.Read()
            if ev_Programs is None or ev_Programs == 'Exit':
                winActive_Programs = False
                win_Programs.Close()
                win_Main.UnHide()
                break

            # QGIS Label Updates Window
            # If Labels window is not active and the event on the programs page is to press 'QGIS Label Updates'...
            if not winActive_Labels and ev_Programs == 'QGIS Label Updates':
                winActive_Labels = True
                win_Programs.Hide()
                # Note: a key (as seen in the layout below in red) allows you to reference this user input in a fxn (see below)
                layout_Labels = [[sg.Text('Please enter the Working Interest Detail and Ownership Labels')],
                                 [sg.Checkbox("Webscrape these files?", key = '-scrape-')],
                                 [sg.Text("Working Interest Detail: "),sg.FileBrowse('Browse', target = '-wi-'), sg.InputText(key='-wi-', size=(65,1))],
                                 [sg.Text("Ownership Labels: "),sg.FileBrowse('Browse', target = '-or-'), sg.InputText(key='-or-', size=(65,1))],
                                 [sg.Text('', key='-output-', size=(10, 1))],
                                 [sg.Submit()]]
                win_Labels = sg.Window('QGIS Label Updates').Layout(layout_Labels)
                while True:
                    ev_Labels, vals_Labels = win_Labels.Read()
                    if ev_Labels is None or ev_Labels == 'Exit':
                        winActive_Labels = False
                        win_Labels.Close()
                        win_Programs.UnHide()
                        break
                    # This is where it starts getting different again.. pay attention.
                    # If the event/click is on Submit:
                    if ev_Labels == 'Submit':
                        # Start a thread (almost like a separate program), that takes the fxn updateLabels, with the
                        # given arguments (using values read pertaining to a specific key), then start it.
                        labelupdate = threading.Thread(target = updateLabels, args = (vals_Labels['-scrape-'],
                                                              vals_Labels['-wi-'], vals_Labels['-or-']), daemon=True)
                        labelupdate.start()
                        # Then update the blank text box in the layout with a message saying it has started.
                        win_Labels.FindElement('-output-').Update(threadMessage)

            if not winActive_Layer and ev_Programs == 'QGIS Layer Updates':
                winActive_Layer = True
                win_Programs.Hide()
                layout_Layer = [[sg.Text('The program to update QGIS Ownership and Leasehold \n'
                                         'has to be run within QGIS. Open up the Python Console \n'
                                         'and type in the command "updateMapLayers()". If this \n'
                                         'doesn\'t work, you can run the script, save it, and try \n'
                                         'again. Then, copy and paste the styles from the old \n'
                                         'layers to the new layers. \n'
                                         'More info in the manual below:')],
                                [sg.Button('Open Job Manual')]]
                win_Layer = sg.Window('QGIS Layer Updates').Layout(layout_Layer)
                while True:
                    ev_Layer, vals_Layer = win_Layer.Read()
                    if ev_Layer is None or ev_Layer == 'Exit':
                        winActive_Layer = False
                        win_Layer.Close()
                        win_Programs.UnHide()
                        break
                    if ev_Layer == "Open Job Manual":
                        os.startfile(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\MapsJOB MANUAL.docx')


            # Metes and Bounds Calculator
            if not winActive_MB and ev_Programs == 'Decimal Degrees M&B Calculator':
                winActive_MB = True
                win_Programs.Hide()
                layout_MB = [[sg.Text('Please enter the Metes and Bounds from the description: ')],
                             [sg.Text("Enter Here: "), sg.Combo(['N','S'], key = 'northing', size = [3,1]), sg.InputText(key = 'degrees', size = [5,1]),
                              sg.InputText(key = 'minutes', size = [5,1]),sg.InputText(key = 'seconds', size = [5,1]),sg.Combo(['E','W'], key = 'easting', size = [3,1]),
                              sg.InputText(key = 'feet', size = [6,1]), sg.Combo(['feet', 'rods', 'chains'], key = 'unit')],
                             [sg.Text('', key = '-output-', size = (50,1))],
                             [sg.Submit()]]
                win_MB = sg.Window('Decimal Degrees M&B Calculator').Layout(layout_MB)
                while True:
                    ev_MB, vals_MB = win_MB.Read()
                    if ev_MB is None or ev_Programs == 'Exit':
                        winActive_MB = False
                        win_MB.Close()
                        win_Programs.UnHide()
                        break
                    if ev_MB == 'Submit':
                        # Here, instead of starting a new instance (or thread), I just use the main thread
                        # because its automatic.. I just run the function as is
                        output = mbDegrees(vals_MB['northing'], vals_MB['degrees'], vals_MB['minutes'], vals_MB['seconds'], vals_MB['easting'], vals_MB['feet'], vals_MB['unit'])
                        win_MB.FindElement('-output-').Update(output)

            # QGIS M&B Import/Calculator
            # Honestly, this was useful before the Ultimate M&B was developed, take it out if you really want to.
            if not winActive_MBImport and ev_Programs == 'QGIS M&B Import/Calculator':
                winActive_MBImport = True
                win_Programs.Hide()
                layout_MBImport = [[sg.Text('Please enter the Metes and Bounds from the description: \n'
                                            'Make sure you enter in the right Northing/Eastings (if not, it will crash)')],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing1', size=[3, 1]), sg.InputText(key='degrees1', size=[5, 1]), sg.InputText(key='minutes1', size=[5, 1]),
                                    sg.InputText(key='seconds1', size=[5, 1]),sg.Combo(['E', 'W'], key='easting1', size=[3, 1]), sg.InputText(key='feet1', size=[6, 1]),sg.Text('', key='-output-1', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing2', size=[3, 1]), sg.InputText(key='degrees2', size=[5, 1]), sg.InputText(key='minutes2', size=[5, 1]),
                                    sg.InputText(key='seconds2', size=[5, 1]), sg.Combo(['E', 'W'], key='easting2', size=[3, 1]),sg.InputText(key='feet2', size=[6, 1]), sg.Text('', key='-output-2', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing3', size=[3, 1]),sg.InputText(key='degrees3', size=[5, 1]), sg.InputText(key='minutes3', size=[5, 1]),
                                    sg.InputText(key='seconds3', size=[5, 1]),sg.Combo(['E', 'W'], key='easting3', size=[3, 1]),sg.InputText(key='feet3', size=[6, 1]), sg.Text('', key='-output-3', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing4', size=[3, 1]),sg.InputText(key='degrees4', size=[5, 1]), sg.InputText(key='minutes4', size=[5, 1]),
                                    sg.InputText(key='seconds4', size=[5, 1]),sg.Combo(['E', 'W'], key='easting4', size=[3, 1]),sg.InputText(key='feet4', size=[6, 1]), sg.Text('', key='-output-4', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing5', size=[3, 1]),sg.InputText(key='degrees5', size=[5, 1]), sg.InputText(key='minutes5', size=[5, 1]),
                                    sg.InputText(key='seconds5', size=[5, 1]),sg.Combo(['E', 'W'], key='easting5', size=[3, 1]),sg.InputText(key='feet5', size=[6, 1]), sg.Text('', key='-output-5', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing6', size=[3, 1]),sg.InputText(key='degrees6', size=[5, 1]), sg.InputText(key='minutes6', size=[5, 1]),
                                    sg.InputText(key='seconds6', size=[5, 1]),sg.Combo(['E', 'W'], key='easting6', size=[3, 1]), sg.InputText(key='feet6', size=[6, 1]), sg.Text('', key='-output-6', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing7', size=[3, 1]),sg.InputText(key='degrees7', size=[5, 1]), sg.InputText(key='minutes7', size=[5, 1]),
                                    sg.InputText(key='seconds7', size=[5, 1]),sg.Combo(['E', 'W'], key='easting7', size=[3, 1]),sg.InputText(key='feet7', size=[6, 1]), sg.Text('', key='-output-7', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing8', size=[3, 1]),sg.InputText(key='degrees8', size=[5, 1]), sg.InputText(key='minutes8', size=[5, 1]),
                                    sg.InputText(key='seconds8', size=[5, 1]),sg.Combo(['E', 'W'], key='easting8', size=[3, 1]),sg.InputText(key='feet8', size=[6, 1]), sg.Text('', key='-output-8', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing9', size=[3, 1]),sg.InputText(key='degrees9', size=[5, 1]), sg.InputText(key='minutes9', size=[5, 1]),
                                    sg.InputText(key='seconds9', size=[5, 1]),sg.Combo(['E', 'W'], key='easting9', size=[3, 1]),sg.InputText(key='feet9', size=[6, 1]), sg.Text('', key='-output-9', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing10', size=[3, 1]),sg.InputText(key='degrees10', size=[5, 1]), sg.InputText(key='minutes10', size=[5, 1]),
                                    sg.InputText(key='seconds10', size=[5, 1]),sg.Combo(['E', 'W'], key='easting10', size=[3, 1]),sg.InputText(key='feet10', size=[6, 1]), sg.Text('', key='-output-10', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing11', size=[3, 1]),sg.InputText(key='degrees11', size=[5, 1]), sg.InputText(key='minutes11', size=[5, 1]),
                                    sg.InputText(key='seconds11', size=[5, 1]),sg.Combo(['E', 'W'], key='easting11', size=[3, 1]),sg.InputText(key='feet11', size=[6, 1]), sg.Text('', key='-output-11', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing12', size=[3, 1]),sg.InputText(key='degrees12', size=[5, 1]), sg.InputText(key='minutes12', size=[5, 1]),
                                    sg.InputText(key='seconds12', size=[5, 1]),sg.Combo(['E', 'W'], key='easting12', size=[3, 1]),sg.InputText(key='feet12', size=[6, 1]), sg.Text('', key='-output-12', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing13', size=[3, 1]),sg.InputText(key='degrees13', size=[5, 1]), sg.InputText(key='minutes13', size=[5, 1]),
                                    sg.InputText(key='seconds13', size=[5, 1]),sg.Combo(['E', 'W'], key='easting13', size=[3, 1]),sg.InputText(key='feet13', size=[6, 1]), sg.Text('', key='-output-13', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing14', size=[3, 1]),sg.InputText(key='degrees14', size=[5, 1]), sg.InputText(key='minutes14', size=[5, 1]),
                                    sg.InputText(key='seconds14', size=[5, 1]),sg.Combo(['E', 'W'], key='easting14', size=[3, 1]),sg.InputText(key='feet14', size=[6, 1]), sg.Text('', key='-output-14', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing15', size=[3, 1]), sg.InputText(key='degrees15', size=[5, 1]), sg.InputText(key='minutes15', size=[5, 1]),
                                    sg.InputText(key='seconds15', size=[5, 1]),sg.Combo(['E', 'W'], key='easting15', size=[3, 1]),sg.InputText(key='feet15', size=[6, 1]), sg.Text('', key='-output-15', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing16', size=[3, 1]),sg.InputText(key='degrees16', size=[5, 1]), sg.InputText(key='minutes16', size=[5, 1]),
                                    sg.InputText(key='seconds16', size=[5, 1]),sg.Combo(['E', 'W'], key='easting16', size=[3, 1]),sg.InputText(key='feet16', size=[6, 1]), sg.Text('', key='-output-16', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing17', size=[3, 1]),sg.InputText(key='degrees17', size=[5, 1]), sg.InputText(key='minutes17', size=[5, 1]),
                                    sg.InputText(key='seconds17', size=[5, 1]),sg.Combo(['E', 'W'], key='easting17', size=[3, 1]),sg.InputText(key='feet17', size=[6, 1]), sg.Text('', key='-output-17', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing18', size=[3, 1]),sg.InputText(key='degrees18', size=[5, 1]), sg.InputText(key='minutes18', size=[5, 1]),
                                    sg.InputText(key='seconds18', size=[5, 1]),sg.Combo(['E', 'W'], key='easting18', size=[3, 1]),sg.InputText(key='feet18', size=[6, 1]), sg.Text('', key='-output-18', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing19', size=[3, 1]),sg.InputText(key='degrees19', size=[5, 1]), sg.InputText(key='minutes19', size=[5, 1]),
                                    sg.InputText(key='seconds19', size=[5, 1]),sg.Combo(['E', 'W'], key='easting19', size=[3, 1]),sg.InputText(key='feet19', size=[6, 1]), sg.Text('', key='-output-19', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing20', size=[3, 1]),sg.InputText(key='degrees20', size=[5, 1]), sg.InputText(key='minutes20', size=[5, 1]),
                                    sg.InputText(key='seconds20', size=[5, 1]),sg.Combo(['E', 'W'], key='easting20', size=[3, 1]),sg.InputText(key='feet20', size=[6, 1]), sg.Text('', key='-output-20', size=(30, 1))],

                                   [sg.Submit()]]
                win_MBImport = sg.Window('QGIS M&B Import/Calculator').Layout(layout_MBImport)
                while True:
                    ev_MBImport, vals_MBImport = win_MBImport.Read()
                    if ev_MBImport is None or ev_MBImport == 'Exit':
                        winActive_MBImport = False
                        win_MBImport.Close()
                        win_Programs.UnHide()
                        break
                    if ev_MBImport == 'Submit':
                        output1 = mbDegreesImport(vals_MBImport['northing1'], vals_MBImport['degrees1'], vals_MBImport['minutes1'],
                                                  vals_MBImport['seconds1'], vals_MBImport['easting1'], vals_MBImport['feet1'])
                        output2 = mbDegreesImport(vals_MBImport['northing2'], vals_MBImport['degrees2'], vals_MBImport['minutes2'],
                                                  vals_MBImport['seconds2'], vals_MBImport['easting2'],vals_MBImport['feet2'])
                        output3 = mbDegreesImport(vals_MBImport['northing3'], vals_MBImport['degrees3'],vals_MBImport['minutes3'],
                                                  vals_MBImport['seconds3'], vals_MBImport['easting3'],vals_MBImport['feet3'])
                        output4 = mbDegreesImport(vals_MBImport['northing4'], vals_MBImport['degrees4'], vals_MBImport['minutes4'],
                                                  vals_MBImport['seconds4'], vals_MBImport['easting4'],  vals_MBImport['feet4'])
                        output5 = mbDegreesImport(vals_MBImport['northing5'], vals_MBImport['degrees5'], vals_MBImport['minutes5'],
                                                  vals_MBImport['seconds5'], vals_MBImport['easting5'],vals_MBImport['feet5'])
                        output6 = mbDegreesImport(vals_MBImport['northing6'], vals_MBImport['degrees6'],vals_MBImport['minutes6'],
                                                  vals_MBImport['seconds6'], vals_MBImport['easting6'],vals_MBImport['feet6'])
                        output7 = mbDegreesImport(vals_MBImport['northing7'], vals_MBImport['degrees7'],vals_MBImport['minutes7'],
                                                  vals_MBImport['seconds7'], vals_MBImport['easting7'],vals_MBImport['feet7'])
                        output8 = mbDegreesImport(vals_MBImport['northing8'], vals_MBImport['degrees8'],vals_MBImport['minutes8'],
                                                  vals_MBImport['seconds8'], vals_MBImport['easting8'],vals_MBImport['feet8'])
                        output9 = mbDegreesImport(vals_MBImport['northing9'], vals_MBImport['degrees9'],vals_MBImport['minutes9'],
                                                  vals_MBImport['seconds9'], vals_MBImport['easting9'],vals_MBImport['feet9'])
                        output10 = mbDegreesImport(vals_MBImport['northing10'], vals_MBImport['degrees10'],vals_MBImport['minutes10'],
                                                  vals_MBImport['seconds10'], vals_MBImport['easting10'],vals_MBImport['feet10'])
                        output11 = mbDegreesImport(vals_MBImport['northing11'], vals_MBImport['degrees11'],vals_MBImport['minutes11'],
                                                  vals_MBImport['seconds11'], vals_MBImport['easting11'], vals_MBImport['feet11'])
                        output12 = mbDegreesImport(vals_MBImport['northing12'], vals_MBImport['degrees12'],vals_MBImport['minutes12'],
                                                  vals_MBImport['seconds12'], vals_MBImport['easting12'],vals_MBImport['feet12'])
                        output13 = mbDegreesImport(vals_MBImport['northing13'], vals_MBImport['degrees13'],vals_MBImport['minutes13'],
                                                  vals_MBImport['seconds13'], vals_MBImport['easting13'],vals_MBImport['feet13'])
                        output14 = mbDegreesImport(vals_MBImport['northing14'], vals_MBImport['degrees14'],vals_MBImport['minutes14'],
                                                  vals_MBImport['seconds14'], vals_MBImport['easting14'],vals_MBImport['feet14'])
                        output15 = mbDegreesImport(vals_MBImport['northing15'], vals_MBImport['degrees15'],vals_MBImport['minutes15'],
                                                  vals_MBImport['seconds15'], vals_MBImport['easting15'], vals_MBImport['feet15'])
                        output16 = mbDegreesImport(vals_MBImport['northing16'], vals_MBImport['degrees16'], vals_MBImport['minutes16'],
                                                  vals_MBImport['seconds16'], vals_MBImport['easting16'],vals_MBImport['feet16'])
                        output17 = mbDegreesImport(vals_MBImport['northing17'], vals_MBImport['degrees17'],vals_MBImport['minutes17'],
                                                  vals_MBImport['seconds17'], vals_MBImport['easting17'],vals_MBImport['feet17'])
                        output18 = mbDegreesImport(vals_MBImport['northing18'], vals_MBImport['degrees18'], vals_MBImport['minutes18'],
                                                  vals_MBImport['seconds18'], vals_MBImport['easting18'], vals_MBImport['feet18'])
                        output19 = mbDegreesImport(vals_MBImport['northing19'], vals_MBImport['degrees19'],vals_MBImport['minutes19'],
                                                  vals_MBImport['seconds19'], vals_MBImport['easting19'],vals_MBImport['feet19'])
                        output20 = mbDegreesImport(vals_MBImport['northing20'], vals_MBImport['degrees20'],vals_MBImport['minutes20'],
                                                   vals_MBImport['seconds20'], vals_MBImport['easting20'],vals_MBImport['feet20'])
                        # Update Outputs ---------
                        win_MBImport.FindElement('-output-1').Update(output1)
                        win_MBImport.FindElement('-output-2').Update(output2)
                        win_MBImport.FindElement('-output-3').Update(output3)
                        win_MBImport.FindElement('-output-4').Update(output4)
                        win_MBImport.FindElement('-output-5').Update(output5)
                        win_MBImport.FindElement('-output-6').Update(output6)
                        win_MBImport.FindElement('-output-7').Update(output7)
                        win_MBImport.FindElement('-output-8').Update(output8)
                        win_MBImport.FindElement('-output-9').Update(output9)
                        win_MBImport.FindElement('-output-10').Update(output10)
                        win_MBImport.FindElement('-output-11').Update(output11)
                        win_MBImport.FindElement('-output-12').Update(output12)
                        win_MBImport.FindElement('-output-13').Update(output13)
                        win_MBImport.FindElement('-output-14').Update(output14)
                        win_MBImport.FindElement('-output-15').Update(output15)
                        win_MBImport.FindElement('-output-16').Update(output16)
                        win_MBImport.FindElement('-output-17').Update(output17)
                        win_MBImport.FindElement('-output-18').Update(output18)
                        win_MBImport.FindElement('-output-19').Update(output19)
                        win_MBImport.FindElement('-output-20').Update(output20)

                        output_total = np.vstack((output1,output2,output3,output4,output5,output6,output7,output8,output9,
                                                     output10,output11,output12,output13,output14,output15,output16,output17,
                                                     output18,output19,output20))

                        output_df = pd.DataFrame(data=output_total, columns = ['Azimuth', 'Distance'])
                        pd.DataFrame.to_csv(self=output_df, path_or_buf=r'C:\Users\Accounting\Downloads\MBImport.csv',
                                            encoding='utf-8', index=False, sep=';')

            if not winActive_MBExtract and ev_Programs == 'M&B String Extractor':
                winActive_MBExtract = True
                win_Programs.Hide()
                layout_MBExtract = [[sg.Text("Enter M&B Description: ")],
                                    [sg.InputText('', key = '-string-', size = [50,30])],
                                    [sg.Text('', key='-output-', size=[50, 30])],
                                    [sg.Text(r"Remember: Don't trust computers.. check your output")],
                                    [sg.Submit()]]
                win_MBExtract = sg.Window('M&B String Extractor').Layout(layout_MBExtract)
                while True:
                    ev_MBExtract, vals_MBExtract = win_MBExtract.Read()
                    if ev_MBExtract is None or ev_Programs == 'Exit':
                        winActive_MBExtract = False
                        win_MBExtract.Close()
                        win_Programs.UnHide()
                        break
                    if ev_MBExtract == 'Submit':
                        output = MBExtract(vals_MBExtract['-string-'])
                        win_MBExtract.FindElement('-output-').Update(output)

            if not winActive_MBUltimate and ev_Programs == 'The Ultimate M&B Program':
                winActive_MBUltimate = True
                win_Programs.Hide()
                layout_MBUltimate = [[sg.Text("""This program does the following:
                                                1. Separates a raw M&B Description and ouputs here
                                                2. Calculates Decimal Degrees from Northing/Easting
                                                3. Exports the conversion as MBUltimate in the Downloads
                                                
                                                IMPORTANT NOTE:
                                                Make sure to quality check the description inputted.
                                                Make sure the northing/easting isn't separated by
                                                spaces, make sure that there aren't any weird
                                                characters, etc. If you don't, the program won't work.""")],
                                     [sg.Text('Enter the description here: '), sg.InputText('', key = '-input-', size = [50, 10])],
                                     [sg.Text('', key='-output-', size = (50,30))],
                                     [sg.Text(r"Remember: Don't trust computers.. check your output")],
                                     [sg.Text('Next Step: Import MBUltimate into the M&B Plugin which will draw it for you.')],
                                     [sg.Submit()]
                                     ]


                win_MBUltimate = sg.Window('The Ultimate M&B Program').Layout(layout_MBUltimate)
                while True:
                    ev_MBUltimate, vals_MBUltimate = win_MBUltimate.Read()
                    if ev_MBUltimate is None or ev_Programs == 'Exit':
                        winActive_MBUltimate = False
                        win_MBUltimate.Close()
                        win_Programs.UnHide()
                        break
                    if ev_MBUltimate == 'Submit':
                        output = MBUltimate(vals_MBUltimate['-input-'])
                        win_MBUltimate.FindElement('-output-').Update(output)

            if not winActive_AddTract and ev_Programs == 'Add Tract':
                winActive_AddTract = True
                win_Programs.Hide()
                layout_AddTract = [[sg.Text("""This program does the following:
                1. Formats tract description
                2. Opens iLandman and the Add Tract page
                3. Fills in all the information based on user input

                IMPORTANT NOTE:
                Make sure to quality check the description inputted.
                """)],
                                   [sg.Text('Enter the tract info here: ')],
                                   [sg.Text('Township: '), sg.InputText('', key='-township-', size=[5, 10]), sg.Text('(4S)')],
                                   [sg.Text('Range: '), sg.InputText('', key='-range-', size=[5, 10]), sg.Text('(3W)')],
                                   [sg.Text('Section: '), sg.InputText('', key='-section-', size=[5, 10]), sg.Text('(29)')],
                                   [sg.Text('Tract Number: '), sg.InputText('', key='-tractnumber-', size=[5, 10]), sg.Text('(5)')],
                                   [sg.Text('County: '), sg.Combo(['Duchesne', 'Uintah'], key='-county-', size=[10, 10]), sg.Text('(Duchesne)')],
                                   [sg.Text('Gross Acreage: '), sg.InputText('', key='-acres-', size=[7, 10]), sg.Text('(100)')],
                                   [sg.Text('Description: '), sg.InputText('', key='-description-', size=[30, 10]), sg.Text('(Section 29:..)')],
                                   [sg.Checkbox('Add to iLandman?', key = '-iLm-'), sg.Checkbox('M&B?', key='-MB-'), sg.Checkbox('Add to Map?', key='-Map-')],
                                   [sg.Text('', key='-output-', size=(15, 1))],
                                   [sg.Text(r"Remember: Don't trust computers.. check your output")],
                                   [sg.Text(
                                         'Note: If the tract is M&B there are a few places\n'
                                         'where user input are required. Check the Application\n'
                                         'Manual for details.')],
                                   [sg.Submit()]
                                    ]

                win_AddTract = sg.Window('Add Tract').Layout(layout_AddTract)
                while True:
                    ev_AddTract, vals_AddTract = win_AddTract.Read()
                    if ev_AddTract is None or ev_Programs == 'Exit':
                        winActive_AddTract = False
                        win_AddTract.Close()
                        win_Programs.UnHide()
                        break
                    if ev_AddTract == 'Submit':
                        if vals_AddTract['-iLm-']:
                            addtract = threading.Thread(target = iLandmanTract, args = (vals_AddTract['-township-'], vals_AddTract['-range-'],
                                                   vals_AddTract['-section-'], vals_AddTract['-tractnumber-'],
                                                   vals_AddTract['-county-'], vals_AddTract['-acres-'],
                                                   vals_AddTract['-description-'], vals_AddTract['-MB-'],
                                                   vals_AddTract['-Map-']), daemon = True)
                            addtract.start()
                            win_AddTract.FindElement('-output-').Update(threadMessage)
                        if vals_AddTract['-MB-'] and vals_AddTract['-Map-']:
                            mapMB = threading.Thread(target = iLandmanTractMap, args = (vals_AddTract['-township-'], vals_AddTract['-range-'],
                                               vals_AddTract['-section-'], vals_AddTract['-tractnumber-'],
                                               vals_AddTract['-county-'], 'NE',
                                               vals_AddTract['-description-']), daemon = True)
                            mapMB.start()
                            win_AddTract.FindElement('-output-').Update(threadMessage)

            if not winActive_Prod and ev_Programs == 'Update Production Spreadsheet':
                winActive_Prod = True
                win_Programs.Hide()
                layout_Prod = [[sg.Text('Press \'Run\' button to output production csv in the \n'
                                        'Downloads folder. Use the output csv to update the\n'
                                        'production spreadsheet in WEM Financial.')],
                             [sg.Text('', key='-output-', size=(50, 1))],
                             [sg.Submit('Run')]]
                win_Prod = sg.Window('Update Production Spreadsheet').Layout(layout_Prod)
                while True:
                    ev_Prod, vals_Prod = win_Prod.Read()
                    if ev_Prod is None or ev_Programs == 'Exit':
                        winActive_Prod = False
                        win_Prod.Close()
                        win_Programs.UnHide()
                        break
                    if ev_Prod == 'Run':
                        output = productionUpdate()
                        win_Prod.FindElement('-output-').Update(output)

    # Files Window
    if not winActive_Files and ev_Main == 'Files':
        winActive_Files = True
        win_Main.Hide()
        layout_Files = [[sg.Text('Which File?')],
                        [sg.Text('Useful Folders/Documents')],
                        [sg.Button('Job Manual')],
                        [sg.Button('Code')],
                        [sg.Button('Important Excel Documents')],
                        [sg.Text('Maps')],
                        [sg.Button('QGIS Map')],
                        [sg.Button('Other Map Layers')],
                        [sg.Text('Financial Database')],
                        [sg.Button('WEM Financial Database 2020')],
                        [sg.Button('Database Spreadsheet')]
                        ]
        win_Files = sg.Window('Files').Layout(layout_Files)

        while True:
            ev_Files, vals_Files = win_Files.Read()
            if ev_Files is None or ev_Files == 'Exit':
                winActive_Files = False
                win_Files.Close()
                win_Main.UnHide()
                break
            if ev_Files == 'Job Manual':
                os.startfile(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\Maps\JOB MANUAL.docx')
            if ev_Files == 'Important Excel Documents':
                os.startfile(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\Maps\QGIS\IMPORTANT EXCEL DOCS\QGIS')
            if ev_Files == 'QGIS Map':
                os.startfile(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\Maps\QGIS\MAIN MAP.qgz')
            if ev_Files == 'Other Map Layers':
                os.startfile(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\Maps\QGIS\Other Map Layers')
            if ev_Files == 'Code':
                os.startfile(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\Maps\Code')
            if ev_Files == 'WEM Financial Database 2020':
                os.startfile(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\WEM Financial\WEM Financial Database 2020.accdb')
            if ev_Files == 'Database Spreadsheet':
                os.startfile(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\WEM Financial\WEM 2020 Finances.xlsx')




# EXPORTING TO APPLICATION -
# cd //WEM-MASTER/Sensitive Data/WEM Uintah/Maps/WEM_Mapping Application
# EXAMPLE COMPILING CODE
# C:\Users\Accounting\Documents\Peter\Application>pyinstaller WEM_Mapping.py --windowed -i WEM_M_Icon.ico
# --hidden-import time --hidden-import PySimpleGUI --hidden-import os --hidden-import rpy2.robjects
# --hidden-import datetime --hidden-import shutil
# First: cd *directory* -- wherever you are putting the application
# Second: See above. Make sure to include:
#           --windowed (makes it run as a window rather than in the cmd)
#           -i iconfile.ico (reads an icon file, make sure its in your cd)
#           --hidden-import *package name* (imports the packages you used, include all of them)
#           --onefile (compresses several files into one application) (DONT DO THIS ONE ANYMORE, MAKES STARTUP SLOW)


# In order to export the application script as an EXE (application) that you don't have to run thru Pycharm, go the the
# command line (cmd) and get to the folder with the scripts in it by pressing "cd Documents" and replace Documents with
# whatever the next folder is that you are going to. Then, with pyinstaller.py in your folder as well as the icon
# converted to an .ico file, run a command that looks something like that above. You can also look at a Word document
# with the command in it titled "pyinstaller command" or something like that. Just follow what is above, such as using
# hidden imports, making sure its windowed, and putting your icon file in there. Its kindof a pain and isn't really
# necessary as the only purpose is to make it so someone without the Python program can run it, so I don't do it often.


# PLEASE DON'T HESITATE TO CONTACT ME.
# I am still willing to help out with whatever bugs come your way, or whatever ideas you have. I really enjoyed coding
# this and encourage you to continue developing it and making it better. I would suggest setting up a github repository
# as I have done on my account and frequently commit and push your changes to the repository. That way, you have a
# backup of your versions of the code and I can easily access it if you need any help.
# Email: pvankatwyk@gmail.com
# Phone: (209) 637-0795

# TODO: Make a list of things that need to be improved (Make a TODO list)