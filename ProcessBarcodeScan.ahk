;Misc subroutines for ID scan program
global versionPBS = 0.02

;########################################################################
;###### DEBUG POPUP WINDOW ##############################################
;########################################################################
DebugOutput(Level, Var1:="", Var2:="", Var3:="", Var4:="", Var5:="", Var6:="", Var7:="", Var8:="", Var9:="", Var10:="") {
    IF (DebugLevel>=Level) {
        MsgBox %Var1%%Var2%%Var3%%Var4%%Var5%%Var6%%Var7%%Var8%%Var9%%Var10%
    }
    ;Level 1 = Subroutines and Major Timers
    ;Level 2 = Variable Definitions
    ;Level 3 = Comm Port Debug & Frequent Timers
}

;########################################################################
;###### PROCESS SCANNED DATA ############################################
;########################################################################
ProcessScan(scanData) {
    DebugOutput(1, "ProcessScan(", scanData, ")")
    SB_SetText("Scan Queue: " ScanLine.Length())
        
    If (EHP.IsBarcodeNumeric(scanData)) {
        DebugOutput(2, "Barcode is Numeric - ", scanData, " & ", PreCmd)
        responseText = % CodeSend(scanData, tripType)
        DebugOutput(2, "CodeSend(%scanData%, %tripType%) Response: ", responseText)
        SB_SetText("Scan Queue: " ScanLine.Length())

        ;Switch based on PreCmd when scanData is numeric
        Switch PreCmd
        {
        Case "EHP":
            GuiControl,2:,DisplayLine1,Checking Pass Status
            GuiControl,2:,DisplayLine2,%tripType%
            GuiControl,2:,DisplayLine3,- ID %scanData% -
            GuiControl,2:,DisplayLine4,Please Wait for Confirmation
            If !EHP.CheckoutPass(scanData, tripType) {
                Return False   
            }
            GuiControl,OSD:,Var1, Hall Pass Obtained
            GuiControl,,DisplayLine,%responseText%
            GuiControl,2:,DisplayLine1,Hall Pass Obtained
            GuiControl,2:,DisplayLine2,Scan Recorded
            GuiControl,2:,DisplayLine3,
            GuiControl,2:,DisplayLine4,%responseText%
            SystemSound("CheckOutBeep")
            SetTimer, tripTimer, 5000
            Return True
            
        Case "ALT":
            GuiControl,,DisplayLine,%responseText%
            GuiControl,2:,DisplayLine1,Alt-Schedule
            GuiControl,2:,DisplayLine2,Scan Recorded
            GuiControl,2:,DisplayLine3,
            GuiControl,2:,DisplayLine4,%responseText%
            SetTimer, tripTimer, 5000
            Return True
            
        Case "DEV":
            GuiControl,,DisplayLine,%responseText%
            GuiControl,2:,DisplayLine1,Scan Recorded
            GuiControl,2:,DisplayLine2,
            GuiControl,2:,DisplayLine3,
            GuiControl,2:,DisplayLine4,%responseText%
            UpdateDeviceArray(True, responseText)
            SystemSound("DeviceBeep")
            SetTimer, tripTimer, 1000
            Return True
            
        Case "ATT":
            SystemSound("ScanBeep")
            Return True
            
        default:
            IF EHPinterface {
                IF !EHP.TogglePass(scanData, tripType) {
                    Return False
                }
            }
        
            GuiControl,Main:,DisplayLine,%responseText%
            GuiControl,2:,DisplayLine1,Scan Recorded
            GuiControl,2:,DisplayLine2,
            GuiControl,2:,DisplayLine3,
            GuiControl,2:,DisplayLine4,%responseText%
            SetTimer, tripTimer, 5000
            Return True
        }
        ;End of Switch for Numeric scanData
    }
    Else
    {
        assess:=RegExReplace(scanData,"_"," ")
        tripType = %assess%
        If (tripType=="SysRst") {
            SysRst()
        }
        
        PreCmd:= SubStr(tripType, 1, 3)
        DebugOutput(2, "PreCmd - ", PreCmd)
        
        ;Switch based on PreCmd when scanData is non-numeric
        Switch PreCmd
        {
        Case "DEV":
            tripType:= StrSplit(tripType,"DEV").2
            DebugOutput(2, "DEV tripType - ", tripType)
            responseText = % CodeSend(1111, "")
            DebugOutput(2, "Response Text (1111): ", responseText)
            
            Switch responseText
            {
            Case "Device Returned":
                GuiControl,,DisplayLine,%responseText%
                GuiControl,2:,DisplayLine1,Scan Recorded
                GuiControl,2:,DisplayLine2,
                GuiControl,2:,DisplayLine3,
                GuiControl,2:,DisplayLine4,%responseText%
                SystemSound("DeviceBeep")
                SetTimer, tripTimer, 1000
                UpdateDeviceArray(False, tripType)
                PreCmd :=
                Return True
            Default: 
                Return False
            }

        Case "EHP":
            tripType:= StrSplit(tripType,"EHP").2
            DebugOutput(2, "tripType - ", tripType)
            DebugOutput(2, "PreCmd - ", PreCmd)
        
        Case "INF":
            tripType:= StrSplit(tripType,"INF").2
            DebugOutput(2, "tripType - ", tripType)
            DebugOutput(2, "PreCmd - ", PreCmd)
        
        Case "ALT":
            tripType:= StrSplit(tripType,"ALT").2
            DebugOutput(2, "tripType - ", tripType)
            DebugOutput(2, "PreCmd - ", PreCmd)

        Case "ATT":
            scanData:= StrSplit(tripType,"ATT").2
            DebugOutput(2, "tripType - ", tripType)
            DebugOutput(2, "PreCmd - ", PreCmd)
            
            SB_SetText("Scan Queue: " ScanLine.Length())
            responseText = % CodeSend(scanData, tripType)
            SystemSound("ScanBeep")
            Return True
            
        default:
            ;Return False ;End of Switch
        }
        
        page = ChromeInst.GetPageByTitle("PBIS Rewards","startswith",1, BoundCallback)     
        pageText:=page.Evaluate("document.body.textContent;")
        DebugOutput(2, "PageText - ", pageText)
        DebugOutput(2, "Barcode is not Numeric. tripType - ", tripType)
        DebugOutput(2, "PreCmd - ", PreCmd)
        ProgressTime:=30
        GuiControl,Main:,tripType,%tripType%
        GuiControl,Main:,ProgressTime, 30
        SetTimer, progressTimer, 1000
        ;SetTimer, tripTimer, 40000
        GuiControl,,DisplayLine,%scanData%
        GuiControl,2:,DisplayLine1,Pass Set
        GuiControl,2:,DisplayLine2,%tripType%
        GuiControl,2:,DisplayLine3, 
        GuiControl,2:,DisplayLine4,Scan ID to Continue
        
        Return True
    }

    DebugOutput(2, "End of ProcessScan: PreCmd - ", PreCmd)
    Return False
}

;########################################################################
;###### SEND HTTP POST TO GOOGLE API ####################################
;########################################################################
CodeSend(scanData, tripType) {
    try{ ; only way to properly protect from an error here
        whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
        whr.Open("POST", WebAppURL, false)
        whr.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
        http1=SIDNO=
        http2=&tripType=
        ;http3=&LargeGUI=
        httpSend=%http1%%scanData%%http2%%tripType%
        DebugOutput(2, "HTTP: ", httpSend)
        whr.Send(httpSend)
        } catch e {
                  SystemSound("FailBeep")
                  Return False
                  }

    SB_SetText("Scan Queue: " ScanLine.Length())
    GuiControl,, scanData,
    Return whr.responseText
}

;########################################################################
;###### UPDATE DEVICE ARRAY #############################################
;########################################################################
UpdateDeviceArray(type, inputText) {
    ;MsgBox %type% - %inputText%
    Switch type
    {
    Case True:
        DeviceLine.Push(inputText)
        ;MsgBox Case True
    Case False:
        ;MsgBox Case False
        Loop % DeviceLine.Length()
        {
            If (InStr(DeviceLine[A_Index], inputText)!=0)
                DeviceLine.RemoveAt(A_Index)
        }
    default:
        ;;;;;;
    }
    ;MsgBox % DeviceLine.Length()
    Loop % DeviceLine.Length()+1
    {
    vDeviceLine := "DeviceLine" A_Index
    Line := DeviceLine[A_Index]
    GuiControl,2:,%vDeviceLine%,%Line%
    ;MsgBox %DeviceLine% - %Line%
    }

    Return
}

;########################################################################
;###### SYSTEM RESET ####################################################
;########################################################################
SysRst() {
    DebugOutput(1, "SysRst()")
    IF (SerialEnable==True) {
        RS232_Close(RS232_FileHandle)
        }
    GuiControl,2:,DisplayLine1,ERROR
    GuiControl,2:,DisplayLine2,NO PASS REQUESTED
    GuiControl,2:,DisplayLine3, 
    GuiControl,2:,DisplayLine4,Reset in 2 Seconds
    SystemSound("FailBeep")
    Sleep 2000
    If EHPinterface {
        try {
            ChromeInst.Kill()
            }
            catch e {
                    SystemSound("FailBeep")
                    }
        ;WinActivate, PBIS Rewards
        ;ControlSend,,{F5},ahk_class Chrome_WidgetWin_1
    }
    Reload
    
}

;########################################################################
;###### CHECK BLUETOOTH CONNECTION ######################################
;########################################################################
CheckBTConnection(BTdeviceName) {
    DllCall("LoadLibrary", "str", "Bthprops.cpl", "ptr")
    VarSetCapacity(BLUETOOTH_DEVICE_SEARCH_PARAMS, 24+A_PtrSize*2, 0)
    NumPut(24+A_PtrSize*2, BLUETOOTH_DEVICE_SEARCH_PARAMS, 0, "uint")
    NumPut(1, BLUETOOTH_DEVICE_SEARCH_PARAMS, 16, "uint")   ; fReturnConnected
    VarSetCapacity(BLUETOOTH_DEVICE_INFO, 560, 0)
    NumPut(560, BLUETOOTH_DEVICE_INFO, 0, "uint")
    Loop ;Looping through connected BT devices and check for deviceName
    {
        GuiControl,Main:,ScanStatus,Checking Connection
        If (A_Index = 1)
        {
            foundedDevice := DllCall("Bthprops.cpl\BluetoothFindFirstDevice", "ptr", &BLUETOOTH_DEVICE_SEARCH_PARAMS, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr")
            if !foundedDevice
                GuiControl,Main:,ScanStatus,No Devices Connected
        }
        else
        {
        if !DllCall("Bthprops.cpl\BluetoothFindNextDevice", "ptr", foundedDevice, "ptr", &BLUETOOTH_DEVICE_INFO)
            GuiControl,Main:,ScanStatus,Device Not Connected
            Break
        }

        If (StrGet(&BLUETOOTH_DEVICE_INFO+64) = BTdeviceName)
        {
            GuiControl,Main:,ScanStatus,Device Connected
            DllCall("Bthprops.cpl\BluetoothFindDeviceClose", "ptr", foundedDevice)
            Return True ;Send true if deviceName is connected
        }
    }
    
DllCall("Bthprops.cpl\BluetoothFindDeviceClose", "ptr", foundedDevice)
Return False ;Send false if BTdeviceName is not connected
}

;########################################################################
;###### AUDIO FILE PLAY #################################################
;########################################################################
SystemSound(soundFile) {

    Switch soundFile
    {
    Case "FailBeep":
        try{
        SoundPlay FailBeep.wav
        } catch e {
            SoundPlay *16
            }
    Case "ScanBeep":
        try{
        SoundPlay ScanBeep.wav
        } catch e {
            SoundPlay *64
            }
    Case "DeviceBeep":
        try{
        SoundPlay DeviceBeep.wav
        } catch e {
            SoundPlay *64
            }
    Case "TripBeep":
        try{
        SoundPlay TripBeep.wav
        } catch e {
            SoundPlay *64
            }
    Case "CheckOutBeep":
        try{
        SoundPlay CheckOutBeep.wav
        } catch e {
            SoundPlay *64
            }
    Case "CheckInBeep":
        try{
        SoundPlay CheckInBeep.wav
        } catch e {
            SoundPlay *64
            }            
    default:
        Return False
    }
Return True
}

;########################################################################
;###### CONFIG FILE CREATION ############################################
;########################################################################
;Create a new configuration file with default values
CfgFileCreate() {
	MsgBox Can't open "IDscanCFG.ini". A new default config file has been created. Close the program and change default values before continuing.
    cfgFile.Close()
	
	IniWrite, Sheets WebApp URL (Scan Log), IDscanCFG.ini, GoogleAPI,       DeploymentURL
    IniWrite, 100 Floor Restroom,           IDscanCFG.ini, GoogleAPI,       DefaultTrip
    IniWrite, False,                        IDscanCFG.ini, HallPassAPI,     EHPinterface
    IniWrite, LoginName in EHP,             IDscanCFG.ini, HallPassAPI,     LoginName
    IniWrite, Password in EHP,              IDscanCFG.ini, HallPassAPI,     LoginPass
    IniWrite, "",                           IDscanCFG.ini, HallPassAPI,     ChromePath
    IniWrite, "",                           IDscanCFG.ini, HallPassAPI,     ProfilePath
    IniWrite, 0,                            IDscanCFG.ini, GUIconfig,       DisplayMonitor
    IniWrite, 0,                            IDscanCFG.ini, GeneralSettings, DebugLevel    
   	IniWrite, Scan Timeout in seconds,      IDscanCFG.ini, GeneralSettings, ScanTimeout
    IniWrite, False,                        IDscanCFG.ini, GeneralSettings, LargeGUI
    IniWrite, H100,                         IDscanCFG.ini, GeneralSettings, LaGUIOffset
    IniWrite, Scan Destination For New Pass,IDscanCFG.ini, GeneralSettings, StartupLine1
    IniWrite, Scan ID to Return Pass,       IDscanCFG.ini, GeneralSettings, StartupLine2
    IniWrite, or,                           IDscanCFG.ini, GeneralSettings, StartupLine3
    IniWrite, Scan ID to Mark as Tardy,     IDscanCFG.ini, GeneralSettings, StartupLine4
    IniWrite, 225,                          IDscanCFG.ini, GeneralSettings, BckGrdA
    IniWrite, W400,                         IDscanCFG.ini, GeneralSettings, BckGrdX
    IniWrite, H300,                         IDscanCFG.ini, GeneralSettings, BckGrdY
    IniWrite, Office,                       IDscanCFG.ini, GeneralSettings, BtnLabel00
    IniWrite, Counselor,                    IDscanCFG.ini, GeneralSettings, BtnLabel01
    IniWrite, Nurse's Office,               IDscanCFG.ini, GeneralSettings, BtnLabel02
    IniWrite, Cafeteria,                    IDscanCFG.ini, GeneralSettings, BtnLabel03
    IniWrite, Gymnasium,                    IDscanCFG.ini, GeneralSettings, BtnLabel04
    IniWrite, Auditorium,                   IDscanCFG.ini, GeneralSettings, BtnLabel05
    IniWrite, 100 Floor Restroom,           IDscanCFG.ini, GeneralSettings, BtnLabel06
    IniWrite, 200 Floor Restroom,           IDscanCFG.ini, GeneralSettings, BtnLabel07
    IniWrite, 300 Floor Restroom,           IDscanCFG.ini, GeneralSettings, BtnLabel08
	IniWrite, False,                        IDscanCFG.ini, CommSetting,     HardwareScan
	IniWrite, COM3,                         IDscanCFG.ini, CommSetting,     RS232Port
	IniWrite, 9600,                         IDscanCFG.ini, CommSetting,     RS232Baud
	IniWrite, N,                            IDscanCFG.ini, CommSetting,     RS232Parity
	IniWrite, 8,                            IDscanCFG.ini, CommSetting,     RS232Data
	IniWrite, 1,                            IDscanCFG.ini, CommSetting,     RS232Stop
	IniWrite, CR,                           IDscanCFG.ini, CommSetting,     BarcodeSuffix
  	IniWrite, False,                        IDscanCFG.ini, CommSetting,     BTenable
    IniWrite, BarCode Scanner spp,          IDscanCFG.ini, CommSetting,     BTdeviceName
	ExitApp 
	Exit Script
}
