;################### IDscan-HallPass.ahk ###################

Global version := "1.0.1"

#SingleInstance Force
#NoEnv
#Include lib\Chrome.ahk\Chrome.ahk ;https://github.com/G33kDude/Chrome.ahk.git
#Include lib\Navigate360-API.ahk\Navigate360-API.ahk
#Include ChromeScan.ahk
#Include ProcessBarcodeScan.ahk
#Include SerialComm.ahk
SetWorkingDir %A_ScriptDir%
SetBatchLines -1

;Check for existing configuration file and generate new if not found
cfgFile := FileOpen("IDscanCFG.ini", "r")
if !IsObject(cfgFile) {
	CfgFileCreate()
    }

;########################################################################
;###### CONFIG FILE READ ################################################
;########################################################################
{ ;Read configuration variables from file and format URLs
IniRead, WebAppURL,    IDscanCFG.ini,       GoogleAPI,          DeploymentURL
IniRead, DefaultTrip,  IDscanCFG.ini,       GoogleAPI,          DefaultTrip
IniRead, EHPinterface, IDscanCFG.ini,       HallPassAPI,        EHPinterface
IniRead, ChromePath,   IDscanCFG.ini,       HallPassAPI,        ChromePath
IniRead, ProfilePath,  IDscanCFG.ini,       HallPassAPI,        ProfilePath
IniRead, LoginName,    IDscanCFG.ini,       HallPassAPI,        LoginName
IniRead, LoginPass,    IDscanCFG.ini,       HallPassAPI,        LoginPass
IniRead, DisplayMonitor,IDscanCFG.ini,      GUIconfig,          DisplayMonitor
IniRead, DebugLevel,   IDscanCFG.ini,       GeneralSettings,    DebugLevel
IniRead, ScanTimeout,  IDscanCFG.ini,       GeneralSettings,    ScanTimeout
IniRead, LargeGUI,     IDscanCFG.ini,       GeneralSettings,    LargeGUI
IniRead, LaGUIOffset,  IDscanCFG.ini,       GeneralSettings,    LaGUIOffset
IniRead, StartupLine1, IDscanCFG.ini,       GeneralSettings,    StartupLine1
IniRead, StartupLine2, IDscanCFG.ini,       GeneralSettings,    StartupLine2
IniRead, StartupLine3, IDscanCFG.ini,       GeneralSettings,    StartupLine3
IniRead, StartupLine4, IDscanCFG.ini,       GeneralSettings,    StartupLine4
IniRead, BckGrdA,      IDscanCFG.ini,       GeneralSettings,    bckgrdA
IniRead, BckGrdX,      IDscanCFG.ini,       GeneralSettings,    bckgrdX
IniRead, BckGrdY,      IDscanCFG.ini,       GeneralSettings,    bckgrdY
IniRead, BtnLabel00,   IDscanCFG.ini,       GeneralSettings,    BtnLabel00
IniRead, BtnLabel01,   IDscanCFG.ini,       GeneralSettings,    BtnLabel01
IniRead, BtnLabel02,   IDscanCFG.ini,       GeneralSettings,    BtnLabel02
IniRead, BtnLabel03,   IDscanCFG.ini,       GeneralSettings,    BtnLabel03
IniRead, BtnLabel04,   IDscanCFG.ini,       GeneralSettings,    BtnLabel04
IniRead, BtnLabel05,   IDscanCFG.ini,       GeneralSettings,    BtnLabel05
IniRead, BtnLabel06,   IDscanCFG.ini,       GeneralSettings,    BtnLabel06
IniRead, BtnLabel07,   IDscanCFG.ini,       GeneralSettings,    BtnLabel07
IniRead, BtnLabel08,   IDscanCFG.ini,       GeneralSettings,    BtnLabel08
IniRead, SerialEnable, IDscanCFG.ini,       CommSetting,        HardwareScan
IniRead, RS232_Port,   IDscanCFG.ini,       CommSetting,        RS232Port
IniRead, RS232_Baud,   IDscanCFG.ini,       CommSetting,        RS232Baud
IniRead, RS232_Parity, IDscanCFG.ini,       CommSetting,        RS232Parity
IniRead, RS232_Data,   IDscanCFG.ini,       CommSetting,        RS232Data
IniRead, RS232_Stop,   IDscanCFG.ini,       CommSetting,        RS232Stop
IniRead, scanSuffix,   IDscanCFG.ini,       CommSetting,        BarcodeSuffix
IniRead, BTenable,     IDscanCFG.ini,       CommSetting,        BTenable
IniRead, BTdeviceName, IDscanCFG.ini,       CommSetting,        BTdeviceName
}

;########################################################################
;########################################################################
;######################## BEGIN MAIN PROGRAM ############################
;########################################################################
;########################################################################

global PreCmd:=
global DebugLevel=DebugLevel
global WebAppURL:=WebAppURL
global LoginName:=LoginName
global LoginPass:=LoginPass
global ProfileURL=ProfileURL
global ScanTimeout:=ScanTimeout*1000
global EHPinterface:=%EHPinterface%
global EHPpin=EHPpin
global KioskCode=KioskCode
global scanID:=
global EventLog:=
global LargeGUI:=%LargeGUI%
global scanData:=
global SerialEnable:=%SerialEnable%
global BTenable:=%BTenable%
global BTdeviceName=BTdeviceName
global queueCount:=0
global ScanLine:=[]
global tripType=DefaultTrip
global bckgrdA=bckgrdA
global bckgrdX=bckgrdX
global bckgrdY=bckgrdY
global GreenColor := "B6D7A8"
global RedColor := "EA9999"
global CustomColor := GreenColor
global DeviceLine := []
global ChromeInst
global ProgressTime:= 30
global CustomDest:=False
#KeyHistory 500

;########################################################################
;###### VERSION AND ENVIRONMENT CONTROL #################################
;########################################################################
{
try {
    oHttp := ComObjCreate("WinHttp.Winhttprequest.5.1")
    httpBase=https://script.google.com/macros/s/AKfycbwlAPQz0HHwVU8qv_xnlVgEcEsdEkVPz-KkdIQ9Q9UStmaIgn-_C8kNvIX9BLtlKlmWVg/exec
    httpsend=%httpBase%?Name=IDscan&Version=%version%
    ;MsgBox %httpsend%
    oHttp.open("GET",httpsend)
    oHttp.send()
    IF (oHttp.responseText!="1") 
    {
        UpdateURL:=oHttp.responseText
        ;MsgBox New version is available. Please download at %msgLink%
    }
} Catch e
{
    MsgBox Error on Version Check: %e%
}

;Check if existing Chrome windows are open
 If (ProcessExist("Chrome.exe") && EHPinterface) {
    MsgBox,,,Please close all Chrome Windows and Try Again., 3
    ;ExitApp
    ;Exit
    ;Process, Close, Chrome.exe
}

;########################################################################
;###### GUI CREATION ####################################################
;########################################################################
IF !LargeGUI {
Gui Main: New, +LabelMainGUI +HwndMainGUI -MaximizeBox
Gui Font, s12, FixedSys
Gui Font
Gui Font, cBlack
Gui Add, Text, hWndhTxt x1024 y0 w2 h768 +0x200
Gui Font, s12 Bold cGreen
If (UpdateURL!="")
{
    Gui Font, s9 cGreen
    Gui Add, Link, x1044 y13 w146 h20 +0x1 +Center, <a href="%UpdateURL%">Download New Version</a>
} Else
    {
        Gui Font, s12 Bold cGreen
        Gui Add, Text, x1034 y10 w156 h30 vScanStatus +Center, Connection Status
    }
Gui Font
Gui Font, s12, FixedSys
Gui Add, Progress, x1034 y39 w156 h32 cGreen vProgressTime Range0-30, 30
Gui Add, Edit, x1033 y43 w158 h24 +Center +ReadOnly +BackgroundTrans +0x200 vtripType, %tripType%
Gui Font

Gui Font
Gui Font, s12 cGreen, FixedSys
Gui Add, GroupBox, x1029 y85 w166 h425 +Center, Other Locations

Gui Font, s12, FixedSys
Gui Add, StatusBar,, Status Bar

Gui Add, Button, hWndhBtn00 gBtnPress00 x1034 y150 w156 h30, %BtnLabel00%
Gui Add, Button, hWndhBtn01 gBtnPress01 x1034 y190 w156 h30, %BtnLabel01%
Gui Add, Button, hWndhBtn02 gBtnPress02 x1034 y230 w156 h30, %BtnLabel02%
Gui Add, Button, hWndhBtn03 gBtnPress03 x1034 y270 w156 h30, %BtnLabel03%
Gui Add, Button, hWndhBtn04 gBtnPress04 x1034 y310 w156 h30, %BtnLabel04%
Gui Add, Button, hWndhBtn05 gBtnPress05 x1034 y350 w156 h30, %BtnLabel05%
Gui Add, Button, hWndhBtn06 gBtnPress06 x1034 y390 w156 h30, %BtnLabel06%
Gui Add, Button, hWndhBtn07 gBtnPress07 x1034 y430 w156 h30, %BtnLabel07%
Gui Add, Button, hWndhBtn08 gBtnPress08 x1034 y470 w156 h30, %BtnLabel08%
;Gui Add, Text, x1024 y80 w176 h2 +0x10
;Gui Font

Gui Font
Gui Font, s12, FixedSys
Gui Font
Gui Font, s10 cBlack, FixedSys
Gui Add, Edit, hWndhLocValue x1034 y110 w156 h30 vManualData +Center, 
Gui Add, Button, Hidden w0 h0 Default gCodeSend, Save
Gui Font
Gui Font, s12 cGreen, FixedSys
Gui Add, GroupBox, x1029 y515 w166 h216 +Center, Pride Pass Request
Gui Add, Button, +Disabled hWndhBtnPridePass x1034 y535 w156 h30, &Print for Date
Gui Add, MonthCal, +Disabled hWndhDate x1030 y570 w163
SB_SetText("v" version)
Gui Show, w1200 h768, ID Scan & Hall Pass
}
else
{
DisplayLine := []
Gui, 2:New, +LastFound +AlwaysOnTop -Caption -MinimizeBox -SysMenu ;+Resize  ; +ToolWindow avoids a taskbar button and an alt-tab menu item.
Gui, Color, %CustomColor%
Gui, Font, s32  ; Set a large font size (32-point).
Gui, Add, Text, +Center %bckgrdX% %LaGUIOffset%
Gui, Add, Text, +Center %bckgrdX% vClock cGreen,%A_Hour%:%A_Min%:%A_Sec%

LoopCount := 0
Loop 4
{
LoopCount += 1
vDisplayLine := "vDisplayLine" LoopCount
;DisplayLine[LoopCount] := _________________________________________
Gui, Add, Text, +Center %bckgrdX% %vDisplayLine% cGreen,_________________________________________
}
Gui, Add, Text, +Center %bckgrdX% 250 cGreen
LoopCount := 0
Loop 6
{
LoopCount += 1
vDeviceLine := "vDeviceLine" LoopCount
;DeviceLine[LoopCount-4] := 
Gui, Add, Text, +Center %bckgrdX% %vDeviceLine% cGreen ;,_________________________________________
}

WinSet, Transparent, %bckgrdA%
Gui, Add, Edit, vScanData x0 y0 w0 h0
Gui, Add, Button, Hidden w0 h0 Default gScanSend, Save
Gui, Show, NoActivate %bckgrdX% %bckgrdY% X0 Y0 ; NoActivate avoids deactivating the currently active window.
GuiControl, 2:, MoveDraw
SetTimer TimeCheck, 1000
}

;########################################################################
;###### OSD for MultiMonitor Displays ###################################
;########################################################################
If (DisplayMonitor != 0)
{
    global CustomColor := "B7E1CD"
    SysGet, Mon, Monitor, %DisplayMonitor% ; 1 is the primairy monitor, 2 the secondairy etc. 
    bckgrdX := "W" Abs(MonRight - MonLeft)
    bckgrdY := "H" Abs(MonTop - MonBottom)
    Xpos := "X" MonLeft
    Ypos := "Y" MonTop + Abs(MonTop - MonBottom) -300
    textYpos := "Y" (Abs(MonTop - MonBottom) - 300)
    ;MsgBox, Left: %MonLeft% -- Top: %MonTop% -- Right: %MonRight% -- Bottom %MonBottom% -- bckgrdX: %bckgrdX% bckgrdY: %bckgrdY% textYpos: %textYpos%.

    ;GuiOSD()
    ;Construct main GUI
    ; Can be any RGB color (it will be made transparent below).
    Gui, OSD:New, +LastFound +AlwaysOnTop +ToolWindow -Caption -MinimizeBox -SysMenu ;+Resize  ; +ToolWindow avoids a taskbar button and an alt-tab menu item.
    Gui, Color, cGreen ;%CustomColor%
    Gui, Font, s64  ; Set a large font size (32-point).
    ;Gui, Add, Text, +Center %bckgrdX% %LargeGUIOffset% cLime
    ;Gui, Add, Text, +Center %bckgrdX% vClock cGreen,%A_Hour%:%A_Min%:%A_Sec%
    Gui, Add, Text, %bckgrdX% vVar1 cLime,_________________________________________
    Gui, Add, Text, %bckgrdX% vVar2 cLime,_________________________________________
    ;Gui, Add, Text, %bckgrdX% vVar3 cLime,_________________________________________
    ;Gui, Add, Text, +Center %bckgrdX% vVar4 cGreen,_________________________________________
    ; Make all pixels of this color transparent and make the text itself translucent (150):
    ;WinSet, TransColor, cGreen 250 ;%CustomColor% 250
    ;WinSet, Transparent, 150

    ;SetTimer, UpdateOSD, 200
    ;Gosub, UpdateOSD  ; Make the first update immediate rather than waiting for the timer.
    ;If LargeGUIOffset=="y0"
    Gui, Show, NoActivate W1000 H300 %Xpos% %Ypos% ; NoActivate avoids deactivating the currently active window.
    ;Else
        ;Gui, Show, NoActivate xCenter %LargeGUIOffset% %bckgrdX% %bckgrdY% ; NoActivate avoids deactivating the currently active window.
    
    GuiControl, OSD:, MoveDraw
    Sleep 1000
    GuiControl, OSD:, Hide
}

;Prep environment if Webpage Hall Pass functions are desired
If EHPinterface {
    ChromePath := ChromePath ? ChromePath : A_ScriptDir "\chrome.exe"
    ProfilePath := ProfilePath ? ProfilePath : A_ScriptDir
    ChromeInst := new Chrome(ProfilePath ,,"--no-first-run --app=https://pbisr.navigate360.com/index.php",ChromePath , 9222)
    ChromeInst.WaitForLoad()

    MainWinID := WinExist("ID Scan & Hall Pass")
    If (!MainWinID) {
        MsgBox, Main GUI not found!
        ExitApp
    }
    
    Gui, 3:New, +HwndGui_Hwnd +LastFound +ToolWindow -Caption -Border
    Gui Color, 0x0080FF
    DllCall("SetParent", "uint", Gui_Hwnd, "uint", MainGUI)
    Gui, 3:Show, ;,Child Button
    DllCall("SetWindowPos", "uint", Gui_Hwnd, "uint", 0, "int", 0, "int", 0, "int", 1024, "int", 745, "uint", 0x0010 | 0x0040) ; SWP_NOACTIVATE | SWP_SHOWWINDOW
    
    SetTitleMatchMode, 2
    WinWait, PBIS Rewards
    TargetHwnd := WinExist()
    DllCall("SetParent", "uint", TargetHwnd, "uint", Gui_Hwnd)
    DllCall("SetWindowPos", "uint", TargetHwnd, "uint", 0, "int", -10, "int", -32, "int", 1044, "int", 785, "uint", 0x0010 | 0x0040) ;SWP_NOACTIVATE | SWP_SHOWWINDOW
    ;Gui Main:Show, w1200 h768, ID Scan & Hall Pass
    
    global BoundCallback := Func("Callback")
    Sleep 500
    ;WinActivate, ID Scan & Hall Pass
    ;ControlSend,,{F11},ahk_id %TargetHwnd%
    Sleep 1500
    ; Enable console events and inject the JS payload
    PageInst := ChromeAttach()
    PageInst.WaitForLoad()
    PageInst.Call("Console.enable")
    
    If !EHP.KioskLogin() {
        SysRst()
    }    
    PageInst.Disconnect()
    SystemSound("ScanBeep")
}
}

;########################################################################
;###### CHECK SERIAL PORT FOR DATA ######################################
;########################################################################
IF (SerialEnable==True) {
    scanState = 0
    scanData =
    If (scanSuffix=="CR" || scanSuffix=="CRLF") {
        scanSuffix:=Chr(13)
    }
    
    Switch BTenable
    {
    Case True:
        GuiControl,Main:,ScanStatus,Scanner Startup
        SetTimer, connectionTimer, 1000
    
    Case False:
        GuiControl,Main:,ScanStatus,Scanner Startup
        RS232_Settings = %RS232_Port%:baud=%RS232_Baud% parity=%RS232_Parity% data=%RS232_Data% stop=%RS232_Stop% dtr=Off
        RS232_FileHandle:=RS232_Initialize(RS232_Settings)
        IF (!RS232_FileHandle) {
            MsgBox Port %RS232_Port% Not Available. Exiting.
            ExitApp
            Exit
        }
        SetTimer, SerialTimer, 500
    }
}

;Main program loop
Loop 
{
    scanLength:=ScanLine.Length()
    DebugOutput(3, "ScanLine: ", scanLength)
    IF (scanLength!=0)
    {
    IF (!ProcessScan(ScanLine[1])) {
        SystemSound("FailBeep")
    }
    ScanLine.RemoveAt(1)
    SB_SetText("Scan Queue: " ScanLine.Length())
    }
    Sleep 1000
}

Return
;########################################################################
;########################################################################
;########################## END MAIN PROGRAM ############################
;########################################################################
;########################################################################


;########################################################################
;###### ON GUI CLOSE ####################################################
;########################################################################
{
MainGuiClose:
2GuiClose:
3GuiClose:
IF (SerialEnable==True)
	RS232_Close(RS232_FileHandle)

If EHPinterface {
    WinActivate, PBIS Rewards
    ControlSend,,^+W,ahk_class Chrome_WidgetWin_1
}
ExitApp
Exit
Return
}

;########################################################################
;###### SERIAL READ TIMER ###############################################
;########################################################################
{
SerialTimer:
;Prevent interruption during execution of this timed thread.
;Critical, On
;0xFF in the line below sets the size of the read buffer.
Read_Data := RS232_Read(RS232_FileHandle,"0xFF",RS232_Bytes_Received)
;Break the timer loop if serial port is closed

;Process the data, if there is any.
If (RS232_Bytes_Received > 0) {
    ;Begin Data to ASCII conversion
    ASCII =
    Read_Data_Num_Bytes := StrLen(Read_Data) / 2 ;RS232_Read() returns 2 characters for each byte

    Loop %Read_Data_Num_Bytes%
    {
        StringLeft, Byte, Read_Data, 2
        StringTrimLeft, Read_Data, Read_Data, 2
        Byte = 0x%Byte%
        Byte := Byte + 0 ;Convert to Decimal       
        ASCII_Chr := Chr(Byte)

        ;Send accumulated characters if suffix character detected
        IF (ASCII_Chr == scanSuffix) {
            ScanLine.Push(scanData)
            SB_SetText("Scan Queue: " ScanLine.Length())
            ;tripType = %DefaultTrip%
            ;SetTimer, tripTimer, Off
            scanData =
            } else  {
                    ;Add new character to existing character string
                    scanData = %scanData%%ASCII_Chr%
                    }

    }

    }
;Critical, Off	
Return
}

;########################################################################
;###### CONNECTION TIMER ################################################
;########################################################################
{
connectionTimer:
DebugOutput(3, "connectionTimer Expired")

IF (BTenable==True) {

    IF CheckBTConnection(BTdeviceName) {
        Return
    } 
    SetTimer, connectionTimer, Off
    RS232_Close(RS232_FileHandle)
    RS232_Settings = %RS232_Port%:baud=%RS232_Baud% parity=%RS232_Parity% data=%RS232_Data% stop=%RS232_Stop% dtr=Off
    RS232_FileHandle:=RS232_Initialize(RS232_Settings)
    
    While (RS232_FileHandle == False)
    {
        Loop, 5
        {
            loopTimer:=6-A_Index
            SB_SetText("Trying again in " loopTimer " seconds")
            Sleep 1000
        }
        SB_SetText("Trying to open serial port")
        RS232_FileHandle:=RS232_Initialize(RS232_Settings)
    
        DebugOutput(3, "FileHandle: ", RS232_FileHandle)
    }
    SetTimer, connectionTimer, %ScanTimeout%
    SetTimer, SerialTimer, 500
    GuiControl,Main:,ScanStatus,Scanner Connected
    SB_SetText("")
    Return
}

SetTimer, connectionTimer, Off  ; i.e. the timer turns itself off here.
RS232_Settings = %RS232_Port%:baud=%RS232_Baud% parity=%RS232_Parity% data=%RS232_Data% stop=%RS232_Stop% dtr=Off
RS232_FileHandle:=RS232_Initialize(RS232_Settings)
IF (!RS232_FileHandle) {
    MsgBox Port %RS232_Port% Not Available. Exiting.
    ExitApp
    Exit
}
GuiControl,Main:,ScanStatus,Scanner Connected
SB_SetText("")
;SetTimer, connectionTimer, %ScanTimeout%
SetTimer, SerialTimer, 500
return
}

;########################################################################
;###### TRIP TIMER ######################################################
;########################################################################
{
tripTimer:
SetTimer, tripTimer, Off  ; i.e. the timer turns itself off here.
GuiControl,Main:,ProgressTime, 30
ProgressTime = 30
SetTimer, progressTimer, Off
tripType=%DefaultTrip%
PreCmd :=
;ControlSend,,{F5},ahk_class Chrome_WidgetWin_1
GuiControl,,Var,
GuiControl,Main:,tripType,%tripType%
GuiControl,2:,DisplayLine1,%StartupLine1%
GuiControl,2:,DisplayLine2,%StartupLine2% 
GuiControl,2:,DisplayLine3,%StartupLine3%
GuiControl,2:,DisplayLine4,%StartupLine4%
GuiControl,2:,ScanData,
GuiControl,OSD:,Var1,
GuiControl,OSD:,Var2,
GuiControl, OSD:, Hide

If EHPinterface {
    Sleep 1000
    PageInst := ChromeAttach()
    PageInst.WaitForLoad()
    PageInst.Call("Console.enable")
    If PageCheck(PageInst, "Log in to your Staff Account")
        EHP.KioskLogin()
    PageInst.Disconnect() 
    WinActivate, IDscanLog
}
return
}

;########################################################################
;###### PROGRESS TIMER ##################################################
;########################################################################
{
progressTimer:
;DebugOutput(1, %ProgressTime%)
;MsgBox %ProgressTime%
;SetTimer, progressTimer, Off  ; i.e. the timer turns itself off here.
;tripType=%DefaultTrip%
IF (ProgressTime > 0) {
    ;MsgBox %ProgressTim%
    ProgressTime:=ProgressTime - 1
    GuiControl,Main:, ProgressTime, %ProgressTime%
    GuiControl,Main:,tripType,%tripType%
} else  {
        GuiControl,Main:, ProgressTime, 30
        ProgressTime = 30
        SetTimer, progressTimer, Off
        SetTimer, tripTimer, 500
        }

return
}

;########################################################################
;###### TIME CHECK ######################################################
;########################################################################
{
TimeCheck:
    Switch
    {
    Case (A_Hour=07 && A_Min=00): CustomColor:=RedColor
    Case (A_Hour=07 && A_Min=15): CustomColor:=GreenColor
    Case (A_Hour=07 && A_Min=50): CustomColor:=RedColor
    Case (A_Hour=08 && A_Min=15): CustomColor:=GreenColor
    Case (A_Hour=08 && A_Min=25): CustomColor:=RedColor
    Case (A_Hour=08 && A_Min=50): CustomColor:=GreenColor
    Case (A_Hour=09 && A_Min=20): CustomColor:=RedColor
    Case (A_Hour=09 && A_Min=45): CustomColor:=GreenColor
    Case (A_Hour=10 && A_Min=15): CustomColor:=RedColor
    Case (A_Hour=10 && A_Min=40): CustomColor:=GreenColor
    Case (A_Hour=11 && A_Min=10): CustomColor:=RedColor
    Case (A_Hour=12 && A_Min=05): CustomColor:=GreenColor
    Case (A_Hour=12 && A_Min=35): CustomColor:=RedColor
    Case (A_Hour=13 && A_Min=00): CustomColor:=GreenColor
    Case (A_Hour=13 && A_Min=30): CustomColor:=RedColor
    Case (A_Hour=13 && A_Min=55): CustomColor:=GreenColor
    Case (A_Hour=14 && A_Min=25): CustomColor:=RedColor
    Default:
        GuiControl,2:,Clock,%A_Hour%:%A_Min%:%A_Sec%
        Gui, 2:Color, %CustomColor%    
    }
    
    SetTimer TimeCheck, 1000
Return
}

;########################################################################
;###### G CODE SEND #####################################################
;########################################################################
{
;GUI call for sending manually-typed data
CodeSend:
	Gui, Submit, NoHide
    ;Global PreCmd = SubStr(tripType, 1, 3)
    CustomDest:=True
    IF (!EHP.IsBarcodeNumeric(ManualData)) {
        ManualData:="EHP" ManualData
    }
        
    If !ProcessScan(ManualData) {
        SysRst()
    }
    ;ProcessScan(ManualData)
	;CodeSend(ManualData, tripType)
	GuiControl,, ManualData,
Return
}

;########################################################################
;###### GUI BUTTON PRESS ################################################
;########################################################################
BtnPress00:
BtnPress01:
BtnPress02:
BtnPress03:
BtnPress04:
BtnPress05:
BtnPress06:
BtnPress07:
BtnPress08:
BtnLabel:= StrSplit(A_ThisLabel,"BtnPress").2
Switch BtnLabel
{
Case 00, 01, 02, 03, 04, 05, 06, 07, 08:
    BtnLabel:= "BtnLabel" BtnLabel
    tripType:=%BtnLabel%
    ProgressTime:=30
    GuiControl,Main:,tripType,%tripType%
    GuiControl,Main:,ProgressTime, 30
    SetTimer, progressTimer, 1000

Default:
}

Return