import pyautogui
import playsound
import clipboard
import time
import sys


### EXCEL PART START ###

## Open Brave Browser
bravebrowser=pyautogui.locateCenterOnScreen('images/bravebrowser.png', confidence=0.9) 
if not bravebrowser:
    print('Error locating Brave browser.')
    playsound('/sound/sad.mp3')
    sys.exit()
else:
    pyautogui.click(bravebrowser)
    time.sleep(0.2)

while True:
## Click on GTIN column
    excel_gtin=pyautogui.locateCenterOnScreen('images/excel_gtin.png', grayscale=True)
    if not excel_gtin:
        print('Error locating GTIN column in Excel. Last GTIN:',clipboard.paste())
        playsound('/sound/sad.mp3')
        break

    pyautogui.click(excel_gtin)
    time.sleep(0.2)

## Click on Filters menu
    excel_filtermenu=pyautogui.locateCenterOnScreen('images/excel_filtermenu.png', grayscale=True)  
    if not excel_filtermenu:
        print('Error locating the filter menu in Excel. Last GTIN:',clipboard.paste())
        playsound('/sound/sad.mp3')
        break

    pyautogui.click(excel_filtermenu)
    time.sleep(0.5)

## Click on ReApply Filters
    excel_refilter=pyautogui.locateCenterOnScreen('images/excel_refilter.png')
    if not excel_refilter:
        print('Error locating the reapply button in the filter menu in Excel. Last GTIN:',clipboard.paste())
        playsound('/sound/sad.mp3')
        break

    pyautogui.click(excel_refilter)
    time.sleep(0.2)

## Click on the first GTIN
    pyautogui.click(excel_gtin)
    time.sleep(0.2)
    pyautogui.press('down')

## Copy the first GTIN
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.2)

### EXCEL PART END ###
# ------------------ #
### MT PART START ###

## Open Remote Desktop
    remotedesktop=pyautogui.locateCenterOnScreen('images/remotedesktop.png', confidence=0.9) 
    if not remotedesktop:
        print('Error locating Remote Desktop. Last GTIN:',clipboard.paste())
        playsound('/sound/sad.mp3')
        break

    pyautogui.click(remotedesktop)
    time.sleep(0.2)

## Remove existing GTIN entry
    mt_removegtin=pyautogui.locateCenterOnScreen('images/mt_removegtin.png') 
    if not mt_removegtin:
        print('Error locating GTIN entry clearing button in MT. Last GTIN:',clipboard.paste())
        playsound('/sound/sad.mp3')
        break

    pyautogui.click(mt_removegtin)

## Move to the GTIN input field
    gtin_input=pyautogui.moveRel(-100, 0, duration=0.3)
    pyautogui.click(gtin_input)
    time.sleep(0.2)

## Paste the GTIN // No hotkey support
    pyautogui.keyDown('ctrl')
    time.sleep(0.2)
    pyautogui.press('v')
    time.sleep(0.1)
    pyautogui.keyUp('ctrl')

## Click on Apply filters
    mt_apply=pyautogui.locateCenterOnScreen('images/mt_apply.png') 
    if not mt_apply:
        print('Error locating Apply Filter button in MT. Last GTIN:',clipboard.paste())
        playsound('/sound/sad.mp3')
        break

    pyautogui.click(mt_apply)
    time.sleep(3.0)

## Click on expanding the GTIN versions
    mt_openmenu=pyautogui.locateCenterOnScreen('images/mt_openmenu.png', grayscale=True) 
    if not mt_openmenu:
        print('Error locating the expanding button of the GTIN in MT. Last GTIN:',clipboard.paste())
        playsound('/sound/sad.mp3')
        break

    pyautogui.click(mt_openmenu)
    time.sleep(3.0)

## Check the GTIN version
    version7=pyautogui.locateCenterOnScreen('images/version7.png')
    if not version7:
        version6=pyautogui.locateCenterOnScreen('images/version6.png')
        if not version6:
            version5=pyautogui.locateCenterOnScreen('images/version5.png')
            if not version5:
                version4=pyautogui.locateCenterOnScreen('images/version4.png')
                if not version4:
                    version3=pyautogui.locateCenterOnScreen('images/version3.png')
                    if not version3:
                        version2=pyautogui.locateCenterOnScreen('images/version2.png')
                        if not version2:
                            version1=pyautogui.locateCenterOnScreen('images/version1.png')
                            if not version1:
                                print('Error locating the GTIN version in MT. Last GTIN:',clipboard.paste())
                                playsound('/sound/sad.mp3')
                                break
                            else:
                                pyautogui.click(version1, button='right')
                        else:
                            pyautogui.click(version2, button='right')
                    else:
                        pyautogui.click(version3, button='right')            
                else:
                    pyautogui.click(version4, button='right')
            else:
                pyautogui.click(version5, button='right')
        else:
            pyautogui.click(version6, button='right')
    else:
        pyautogui.click(version7, button='right')

    time.sleep(1.0)

## Select creating a new version of the MasterData
    mt_masterdatamenu=pyautogui.locateCenterOnScreen('images/mt_masterdatamenu.png') 
    if not mt_masterdatamenu:
        print('Error locating the MasterData menu in MT. Last GTIN:',clipboard.paste())
        playsound('/sound/sad.mp3')
        break

    pyautogui.click(mt_masterdatamenu)
    time.sleep(1.0)

    mt_newversion=pyautogui.locateCenterOnScreen('images/mt_newversion.png') 
    if not mt_newversion:
        print('Error locating the MasterData new version button in MT. Last GTIN:',clipboard.paste())
        playsound('/sound/sad.mp3')
        break

    pyautogui.click(mt_newversion)
    time.sleep(3.0)
    
## Click on the MAH selector button
    mt_mah=pyautogui.locateCenterOnScreen('images/mt_mah.png') 
    if not mt_mah:
        print('Error locating the MAH selector button in MT. Last GTIN:',clipboard.paste())
        playsound('/sound/sad.mp3')
        break

    pyautogui.click(mt_mah)
    time.sleep(0.2)

## Select the new MAH // MISSING

## Wait for the EMVO Answer // MISSING



## Minimize the Remote Desktop window
    rd_minimize=pyautogui.locateCenterOnScreen('images/rd_minimize.png', confidence=0.9) 
    if not rd_minimize:
        print('Error locating the button to minimize the Remote Desktop window. Last GTIN:',clipboard.paste())
        playsound('/sound/sad.mp3')
        break

    pyautogui.click(rd_minimize)
    time.sleep(0.2)

## Write x in the Excel
    excel_x=pyautogui.locateCenterOnScreen('images/excel_x.png', grayscale=True)
    
    if not excel_x:
        print('Error locating Update column in Excel. Last GTIN:',clipboard.paste())
        playsound('/sound/sad.mp3')
        break
        
## If successfull maximum 30 seconds
    pyautogui.click(excel_x)
    time.sleep(0.1)
    pyautogui.press('down', interval=0.1)
    pyautogui.typewrite('x\n')


## If not successfull maximum 30 seconds // MISSING
## Write timeout in the Excel // MISSING


## Finish
    print('GTIN done:',clipboard.paste())
    break
