import os
import aqt
from aqt import mw
from aqt.utils import showInfo, qconnect
from anki.importing import TextImporter
from aqt.qt import *
from aqt.utils import showInfo

## Import list from an old addon I wanted to make

def conditionalLoad():
    if checkFolder():
        #checks for folder

        importOnLoad()
        #if present the imports
    else:
        showInfo("There was nothing to import")
        #if not then shows text

def checkFolder():
    absolute_path = os.path.abspath(__file__)
    #gets the current file path

    Folder = "Library/Application Support/Anki2/addons21/123_321/__init__.py"
    #String to remove form file path. Remember to change the 123_321 as this was my jokey folder name

    Folder = absolute_path.replace(Folder, "")
    #Remove the string from the folder path to get the user directory

    projecttestfolder = os.path.join(Folder, ".anki-PPT")
    #create variable to test for .anki folder presence

    if not os.path.exists(projecttestfolder):
        #if not present

        os.mkdir(projecttestfolder)
        #creates folder if not present
        return False
        #returns False i.e. the folder does not exist

    else:
        return True
        #returns true i.e. it does


def importOnLoad():
    #Checks existance of .anki folder

    deck_id = mw.col.decks.id("PowerPoint_Import")
    #Imports cards into PowerPoint_Import deck - change this if you want

    absolute_path = os.path.abspath(__file__)
    #gets file of __init__.py

    addonFile = "Library/Application Support/Anki2/addons21/123_321/__init__.py"
    #truncates file path using method form checkFolder()

    Folder = absolute_path.replace(addonFile, ".anki-PPT/")
    #trucates

    list = os.listdir(Folder) 
    #lists the files/folders within .anki
    
    number_files = len(list)
    #converts the list to a count
    if (number_files == 0): 
        # Checks there are any folders

        showInfo("There was nothing to import")
        #Shows there was nothing to import

        return 0
        # Leaves if there arent 
    i = 1
    #sets count to be 1 as slides start at 1 not 0 in PowerPoint, hence files start at 1 as well
    while True:
        # Loop

        if i == number_files:
            #Checks if there are anyfiles left

            showInfo("Finished")
            break
            #Ends loop
        
        file_1 =  Folder + str(i) + ".csv"
        #creates new filepath

        file = file_1.encode(encoding = 'UTF-8', errors = 'strict')
        #encodes filepath in uniode for textImport

        mw.col.decks.select(deck_id)
        #selects deck
       
        notetype = mw.col.models.by_name("Basic")
        #selects note type
        deck = mw.col.decks.get(deck_id)
        deck['mid'] = notetype['id']
        mw.col.decks.save(deck)
        # and puts cards in the last deck used by the note type
        mw.col.set_aux_notetype_config(
            notetype["id"], "lastDeck", deck_id
        )
        # import into the collection
        ti = TextImporter(mw.col, file)
        ti.initMapping()
        ti.run()
        #Imports file

        os.remove(file_1)
        #removes file
        i = i +1
        #increases loop counter

action = QAction(aqt.mw)
action.setText("Import from PowerPoint")
aqt.mw.form.menuTools.addAction(action)
action.triggered.connect(conditionalLoad)
