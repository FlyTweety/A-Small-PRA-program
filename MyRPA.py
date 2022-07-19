import pyautogui
import time
import xlrd
import pyperclip

#mouseclick event
def mouseClick(clickTimes,lOrR,img,reTry):
    if reTry == 1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                break
            print("No matching image found, retry after 0.1 second")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                print("repeat")
                i += 1
            time.sleep(0.1)

# check data
# cmdType.value  1.0:left-click|2.0:double-left click|3.0:right-click|4.0:input|5.0:wait|6.0:scroll|7.0:branch-jump|8:0:jump
# ctype          0:empty|1:string|2:number|3:data|4:bool|5:error
def dataCheck(sheet):
    checkCmd = True
    #check if empty
    if sheet.nrows<2:
        print("it is empty")
        checkCmd = False
    #check everyrow
    i = 1
    while i < sheet.nrows:
        # col_1 check operation type
        cmdType = sheet.row(i)[0]
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0 
        and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0 and cmdType.value != 7.0):
            print('row ',i+1,", col 1 is error")
            checkCmd = False
        # col_2 check content
        cmdValue = sheet.row(i)[1]
        # click/branch-jump operation, content must be string(path)
        if cmdType.value ==1.0 or cmdType.value == 2.0 or cmdType.value == 3.0 or cmdType.value ==7.0:
            if cmdValue.ctype != 1:
                print('row ',i+1,", col 2 is error")
                checkCmd = False
        # input operation, content cannot be empty
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                print('row ',i+1,", col 2 is error")
                checkCmd = False
        # wait/scroll/jump operation, content must be number
        if cmdType.value == 5.0 or cmdType.value == 6.0 or cmdType.value == 8.0:
            if cmdValue.ctype != 2:
                print('row ',i+1,", col 2 is error")
                checkCmd = False
        i += 1
    return checkCmd

#if target img can be found, return True(goto branch)
def isBranch(sheetNumber, i):
    img = sheets[sheetNumber].row(i)[1].value
    location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
    if location is not None:
        return True
    return False

#main action
def mainWork(cmdSheetNumber):
    i = 1
    while i < sheets[cmdSheetNumber].nrows:
        #get operation type
        cmdType = sheets[cmdSheetNumber].row(i)[0]
        #1 left-click
        if cmdType.value == 1.0:
            #get picture name
            img = sheets[cmdSheetNumber].row(i)[1].value
            reTry = 1
            if sheets[cmdSheetNumber].row(i)[2].ctype == 2 and sheets[cmdSheetNumber].row(i)[2].value != 0:
                reTry = sheets[cmdSheetNumber].row(i)[2].value
            mouseClick(1,"left",img,reTry)
            print("left-click",img)
        #2 double left-click
        elif cmdType.value == 2.0:
            #get picture name
            img = sheets[cmdSheetNumber].row(i)[1].value
            #get repeat times
            reTry = 1
            if sheets[cmdSheetNumber].row(i)[2].ctype == 2 and sheets[cmdSheetNumber].row(i)[2].value != 0:
                reTry = sheets[cmdSheetNumber].row(i)[2].value
            mouseClick(2,"left",img,reTry)
            print("double left-click",img)
        #3 right-click
        elif cmdType.value == 3.0:
            #get picture name
            img = sheets[cmdSheetNumber].row(i)[1].value
            #get repeat times
            reTry = 1
            if sheets[cmdSheetNumber].row(i)[2].ctype == 2 and sheets[cmdSheetNumber].row(i)[2].value != 0:
                reTry = sheets[cmdSheetNumber].row(i)[2].value
            mouseClick(1,"right",img,reTry)
            print("right-click",img) 
        #4 input
        elif cmdType.value == 4.0:
            inputValue = sheets[cmdSheetNumber].row(i)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl','v')
            time.sleep(0.5)
            print("input:",inputValue)                                        
        #5 wait
        elif cmdType.value == 5.0:
            waitTime = sheets[cmdSheetNumber].row(i)[1].value
            time.sleep(waitTime)
            print("wait",waitTime,"秒")
        #6 mouse scroll
        elif cmdType.value == 6.0:
            scroll = sheets[cmdSheetNumber].row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("mouse scroll",int(scroll),"distance")           
        #7 branch-jump
        elif cmdType.value == 7.0:
            if(isBranch(cmdSheetNumber, i)):
                branchSheetNumber = int(sheets[cmdSheetNumber].row(i)[3].value)
                print("Branch taken, go to sheet ", branchSheetNumber)
                mainWork(branchSheetNumber-1)
            else:
                print("Branch not taken")
        #8 jump
        elif cmdType.value == 8.0:
            branchSheetNumber = int(sheets[cmdSheetNumber].row(i)[3].value)
            print("jump to sheet ", branchSheetNumber)
            mainWork(branchSheetNumber-1)
        i += 1
    return 0

# main function 
if __name__ == '__main__':
    print('Welcome to use FlyTweety\'s RPA')
    #open file
    file = 'C:\\Users\\VitoZCY\\Desktop\\MyRPA\\mycmd.xls'
    wb = xlrd.open_workbook(filename=file)
    #get sheets
    sheets = wb.sheets()
    #record which sheet we are now
    sheetNumber = 0
    #check every cmd sheets
    checkCmd = True
    for i in range(len(sheets)):
        checkCmd = dataCheck(sheets[i])
    #start process
    if checkCmd:
        key=input('choose model: 1.do it once 2.never stop \n')
        if key=='1':
            mainWork(0)
        elif key=='2':
            while True:
                mainWork(0)
                time.sleep(0.1)
                print("wait 0.1 second")    
    else:
        print('error or quit!')