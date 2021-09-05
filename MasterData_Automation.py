import pyautogui
import time

def copy():
    pyautogui.hotkey('ctrl', 'c')

def selectall():
    pyautogui.keyDown('ctrl')
    time.sleep(0.1)
    pyautogui.press('a')
    time.sleep(0.1)
    pyautogui.keyUp('ctrl')

def backspace():
    pyautogui.press('backspace')

def paste_NoHotkeySupport():
    pyautogui.keyDown('ctrl')
    time.sleep(0.1)
    pyautogui.press('v')
    time.sleep(0.1)
    pyautogui.keyUp('ctrl')

def write_ok():
    pyautogui.typewrite('x\n',interval=0.05)

def write_issue():
    pyautogui.typewrite('ISSUE',interval=0.05)

def write_issue_description_masterdata_timeout():
    pyautogui.typewrite('MasterData TimeOut - Manual ReCheck\n',interval=0.05)

def click_brave_browser():
    click_on_brave=pyautogui.moveTo(895, 1053, duration=0.3)
    pyautogui.click(click_on_brave)

def click_remote_desktop():
    click_on_rd=pyautogui.moveTo(1204, 1053, duration=0.3)
    pyautogui.click(click_on_rd)

def click_minimize_remote_desktop():
    click_on_minimize_rd=pyautogui.moveTo(1160, 10, duration=0.3)
    pyautogui.click(click_on_minimize_rd)
    
def click_excel_gtin():  
    click_on_gtin=pyautogui.moveTo(140, 310, duration=0.3)
    pyautogui.click(click_on_gtin)

def click_excel_progress():
    click_on_progress=pyautogui.moveTo(310, 310, duration=0.3)
    pyautogui.click(click_on_progress)

def click_excel_description():  
    click_on_description=pyautogui.moveTo(460, 310, duration=0.3)
    pyautogui.click(click_on_description)

def click_set_filter_again():
    click_on_filter=pyautogui.moveTo(1710, 200, duration=0.3)
    pyautogui.click(click_on_filter)
    click_on_filter_again=pyautogui.moveTo(1710, 445, duration=0.3)
    pyautogui.click(click_on_filter_again)

def click_filter_productcode():  
    click_on_filter_productcode=pyautogui.moveTo(1815, 232, duration=0.3)
    pyautogui.click(click_on_filter_productcode)

def click_apply_filter():  
    click_on_apply_filter=pyautogui.moveTo(1850, 960, duration=0.3)
    pyautogui.click(click_on_apply_filter)

def click_masterdata():
    click_on_masterdata=pyautogui.moveTo(120, 215, duration=3)
    pyautogui.click(click_on_masterdata)
    
def select_masterdata():
    click_on_select_masterdata=pyautogui.moveRel(115, 100, duration=0.3)
    pyautogui.click(click_on_select_masterdata)

def select_masterdata_new_version():
    click_on_select_masterdata_new_version=pyautogui.moveRel(160, 25, duration=0.3)
    pyautogui.click(click_on_select_masterdata_new_version)

def open_masterdata():
    click_on_viewmenu=pyautogui.moveTo(133, 33, duration=0.3)
    pyautogui.click(click_on_viewmenu)
    click_on_expand=pyautogui.moveTo(170, 55, duration=0.3)
    pyautogui.click(click_on_expand)

def ok_button():
    click_on_okbutton=pyautogui.moveTo(1320, 950, duration=0.5)
    pyautogui.click(click_on_okbutton)

def select_mah():
    mah_image=pyautogui.locateCenterOnScreen('images/mah.png', confidence=0.9)
    
    if not mah_image:
        click_minimize_remote_desktop()
        click_excel_progress()
        write_issue()
        click_excel_description()
        pyautogui.typewrite('MAH image not found - Manual ReCheck\n',interval=0.05)
    else:
        pyautogui.click(mah_image)
        # MISSING SELECT MAH AND CLICK ON OK

def check_masterdata_version():
    version7=pyautogui.locateCenterOnScreen('images/version7.png', confidence=0.9)

    if not version7:
        print('Ver7 not found')
        version6=pyautogui.locateCenterOnScreen('images/version6.png', confidence=0.9)
        if not version6:
            print('Ver6 not found')
            version5=pyautogui.locateCenterOnScreen('images/version5.png', confidence=0.9)
            if not version5:
                print('Ver5 not found')
                version4=pyautogui.locateCenterOnScreen('images/version4.png', confidence=0.9)
                if not version4:
                    print('Ver4 not found')
                    version3=pyautogui.locateCenterOnScreen('images/version3.png', confidence=0.9)
                    if not version3:
                        print('Ver3 not found')
                        version2=pyautogui.locateCenterOnScreen('images/version2.png', confidence=0.9)
                        if not version2:
                            print('Ver2 not found')
                            version1=pyautogui.locateCenterOnScreen('images/version1.png', confidence=0.9)
                            if not version1:
                                click_minimize_remote_desktop()
                                click_brave_browser()
                                click_excel_progress()
                                write_issue()
                                click_excel_description()
                                pyautogui.typewrite('Version not found - Manual ReCheck\n',interval=0.05)
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

click_brave_browser()
click_excel_progress()
click_set_filter_again()
click_excel_gtin()
copy()
click_remote_desktop()
click_filter_productcode()
selectall()
backspace()
paste_NoHotkeySupport()
click_apply_filter()
click_masterdata()
open_masterdata()
time.sleep(0.5)
check_masterdata_version()
time.sleep(0.5)
select_masterdata()
time.sleep(0.5)
select_masterdata_new_version()
time.sleep(3.0)
select_mah()
