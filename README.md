# A-Small-PRA-program
Help you do the basic routine work, like check in to the game or automatically reply messages.

1. This small project is base on the work of "不高兴就喝水"(you can find him on www.bilibili.com). I made some important changes to the origin version, which makes it much more "smart"
2. You're supposed to use python3.4 or above
3. These packages are required:  
   a.pip install pyperclip
   b.pip install xlrd
   c.pip install pyautogui
   d.pip install opencv-python
   e.pip install pillow
4. Save the ICONS and screenshots for each step to this folder in PNG format (note that if there are multiple same ICONS on the same screen, the one on the upper left will be found by default, so how to capture the screenshot and how large the area is important. For example, it is not possible to capture only the blank part in the middle of the input box
5. In mycmd.xls, configure instructions for each step, such as instruction type 1234 for screen shot file name (English only), instruction 5 for wait time (unit of seconds) and instruction 6 for roller roll distance, with positive numbers indicating roll up and negative numbers indicating roll down.  you can try 200 and minus 200 first
6. Ctrl&C can stop it
