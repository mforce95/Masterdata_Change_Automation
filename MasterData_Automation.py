import pyautogui
from playsound import playsound
import clipboard
import time
import sys

def click_on(filename, gray=False, trust=1.0, mousebutton='left'):
    newfile=pyautogui.locateCenterOnScreen(filename, grayscale=gray, confidence=trust)
    if not newfile:
        print('Error locating ' + filename + ' Last GTIN: ' + clipboard.paste())
        playsound('sound/sad.mp3')
        return False 
    pyautogui.click(newfile, button=mousebutton)
    time.sleep(0.5)
    return True

def read_version(versionfile):
    readversion=pyautogui.locateCenterOnScreen(versionfile, confidence=0.9)
    if readversion:
        pyautogui.click(readversion, button='right')
        return True
    return False

    ### EXCEL PART START ###

    ## Open Brave Browser
click_on(filename='images/bravebrowser.png', trust=0.9)

while True:
    ## Click on GTIN column
    if not click_on(filename='images/excel_gtin.png', gray=True, trust=0.9):
        break

    ## Click on Filters menu
    if not click_on(filename='images/excel_filtermenu.png', gray=True, trust=0.9):
        break

    ## Click on ReApply Filters
    if not click_on(filename='images/excel_refilter.png', trust=0.9):
        break
    time.sleep(1.0)

    ## Click on the first GTIN
    if not click_on(filename='images/excel_gtin.png', gray=True, trust=0.9):
        break
    
    pyautogui.press('down')
    time.sleep(0.2)

    ## Copy the first GTIN
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.2)

    ### EXCEL PART END ###
    # ------------------ #
    ### MT PART START ###

    ## Open Remote Desktop
    if not click_on(filename='images/remotedesktop.png', trust=0.9):
        break
    time.sleep(0.5)

    ## Remove existing GTIN entry
    if not click_on(filename='images/mt_removegtin.png', trust=0.9):
        break

    ## Move to the GTIN input field
    pyautogui.click(pyautogui.moveRel(-100, 0, duration=0.5))

    ## Paste the GTIN // No hotkey support
    pyautogui.keyDown('ctrl')
    time.sleep(0.2)
    pyautogui.press('v')
    time.sleep(0.2)
    pyautogui.keyUp('ctrl')
    time.sleep(0.5)

    ## Click on Apply filters
    if not click_on(filename='images/mt_apply.png', trust=0.9):
        break
    time.sleep(0.5)

    ## Click on expanding the GTIN versions
    if not click_on(filename='images/mt_openmenu.png', gray=True, trust=0.9):
        break
    time.sleep(2.5)

    ## Check the GTIN version
    versions = ['images/version7.png', 'images/version6.png', 'images/version5.png', 'images/version4.png', 'images/version3.png', 'images/version2.png', 'images/version1.png']
    for x in versions:
        if read_version(x):
            break
    time.sleep(0.3)

    ## Select creating a new version of the MasterData
    if not click_on(filename='images/mt_masterdatamenu.png', trust=0.9):
        break
    if not click_on(filename='images/mt_newversion.png', trust=0.9):
        break
    time.sleep(1.0)

    ## Identify German Masterdata - double-check
    itisde = False
    for i in range(2):
        if pyautogui.locateCenterOnScreen('images/mt_isit_de.png', confidence=0.9):
            itisde = True
            break
        time.sleep(1.0)

    ## Click on the MAH selector button
    if not click_on(filename='images/mt_mah.png', trust=0.9):
        break
    time.sleep(0.5)

    ## Select the new MAH
    if itisde:
        if not click_on(filename='images/mt_eid.png', trust=0.9):
            break
    else:
        if not click_on(filename='images/mt_mid.png', trust=0.9):
            break

    ## Click OK 2x
    for i in range(2):
        if not click_on(filename='images/mt_ok.png', trust=0.9):
            break
        time.sleep(0.2)
    time.sleep(1.0)

    ## Select the new upload
    if not click_on(filename='images/mt_regnew.png', trust=0.9, mousebutton='right'):
        break

    ## Upload to EMVO
    if not click_on(filename='images/mt_emvomenu.png', trust=0.9):
        break
    if not click_on(filename='images/mt_emvoupload.png', trust=0.9):
        break
    time.sleep(5.0)

    ## Wait for the EMVO Answer
    uploadsuccess = False
    for n in range(25):
        if pyautogui.locateCenterOnScreen('images/mt_distributed.png', confidence=0.95):
            uploadsuccess = True
            break
        time.sleep(1.0)

    ### MT PART END ###
    # ------------------ #
    ### EXCEL PART START ###

    ## Minimize the Remote Desktop window
    if not click_on(filename='images/rd_minimize.png', trust=0.9):
        break

    ## Locating the 'Update (x)' column in Excel
    if not click_on(filename='images/excel_x.png', gray=True, trust=0.9):
        break
    pyautogui.press('down')
    time.sleep(0.2)

    ## Write 'x' or 'timeout' in the Excel 
    if uploadsuccess:
        pyautogui.typewrite('x\n')
    else:
        pyautogui.typewrite('timeout\n')

    ### EXCEL PART END ###

    ## Finish
    print('GTIN done:',clipboard.paste())
