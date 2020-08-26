
def sound(name):
    from playsound import playsound
    # FXN: Plays a sound
    # INPUT: Name of the file in Sounds folder (see Filepath) -- "beep"
    name = name+'.mp3'
    song = r"\\WEM-MASTER\Working Projects\WEMU Leasing\Python Codes\Python Code\Sounds" + "\\" + name
    playsound(song)