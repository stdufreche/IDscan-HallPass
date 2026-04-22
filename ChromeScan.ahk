;Chrome subroutines for ID scan program
global versionCS = 2

ProcessExist(Name){
    Process, Exist, %Name%
    Return ErrorLevel
}

ChromeAttach() {
    DebugOutput(2, "ChromeAttach()")
    ; Bind the page number to the function for extra information in the callback
    ; Get an instance of the page, passing in the callback function
    while (!IsObject(PageInst)) {
    PageInst := ChromeInst.GetPageByTitle("PBIS Rewards","startswith",1, BoundCallback)
    If A_Index > 30
        Break
    Sleep 500
    }
    
    if !(PageInst)
    {
        MsgBox, Could not retrieve page 
        ChromeInst.Kill()
        ExitApp
    }
    ;PageInstances.Push(PageInst)
    Return PageInst
}

PageCheck(PageInst, Text) {
    DebugOutput(2, "PageCheck(PageInst,", Text)
    EventLog :=
    JSeval =
        (
        (function(){
        const text1 = '%Text%';
            if (document.body.textContent.includes(text1)) {
                console.log("IDscan:" + text1);
            };
        })();
        )
    
    PageInst.Evaluate(JSeval, 5)
    Sleep 250
    DebugOutput(2, "EventLog: ",EventLog)

    If (EventLog == Text) { 
        Return True
    } else  {
            Return False
            }
}

ButtonCheck(PageInst, buttonID, Text) {
    DebugOutput(1, "ButtonCheck(PageInst, ", buttonID, Text, ")")
    EventLog :=
    JSeval =
        (
        (function(){
        const text1 = '%Text%';
        const buttonList = document.querySelectorAll('%buttonID%');
        let present = 0;
        for (let i=0; i<buttonList.length; i++) {
            if (buttonList[i].textContent.includes(text1)) {
                console.log("IDscan:" + i);
                present = 1;
                break;
            };
        }
        if (present==0) {
            console.log("IDscan:" + 999);
            }
        })();
        )
    
    PageInst.Evaluate(JSeval, 5)
    Sleep 250
    DebugOutput(2, "ButtonCheck EventLog: ", EventLog)
    Return EventLog
}

CheckButtonByID(PageInst, buttonID, Text) {
    DebugOutput(2, "CheckButtonByID(PageInst, ", buttonID, Text)
    EventLog :=
    JSeval =
        (
        (function(){
        const text1 = '%Text%';
        const buttonText = document.getElementById('%buttonID%').textContent;
        console.log("IDscan:" + buttonText);
        })();
        )
    
    PageInst.Evaluate(JSeval, 5)
    Sleep 250
    DebugOutput(2, "EventLog: ", EventLog)
    If (EventLog == Text) {
        Return True
        } else {
                Return False
               }

}

Callback(Event) {
    ; Filter for console messages starting with "IDscan:"
    if (Event.Method == "Console.messageAdded"
        && InStr(Event.params.message.text, "IDscan:") == 1)
    {
        ; Strip out the leading AHK:
        EventLog := SubStr(Event.params.message.text, 8)
        DebugOutput(2, "Callback(", EventLog, ")")
    }
}
