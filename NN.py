import pyautogui as pg
import time
import random
import cv2
import keyboard  # import the keyboard library

def main():
    # DETECT IMAGE ON SCREEN
    location = pg.locateCenterOnScreen("pic.jpg", confidence=0.8)
    if location is not None:
        pg.moveTo(location)    
    pg.moveTo(822, 965, 0)
    pg.click()
    pg.typewrite("Test")
    pg.press('enter')
    pg.moveRel(0, -45)  # Move the cursor up by 45 pixels

if __name__ == '__main__':
    for i in range(10):
        main()
        if keyboard.is_pressed('q'):  # if key 'q' is pressed 
            print('You Pressed A Key!')
            break  # finish the loop
    