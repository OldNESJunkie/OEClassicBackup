;*********************************
;**  OE Classic Database Backup **
;**              by             **
;**   Daniel Ford 02/12/2021    **
;**     Updated 01/06/2024      **
;*********************************


;{ Include external libraries
XIncludeFile "Includes\MultiLang.pb"
XIncludeFile "Includes\Registry.pbi"
UseModule Registry
;}

;{ Re-create systray icon if explorer crashes
#TaskbarCreated = #PB_Event_FirstCustomValue
Declare.s SizeIt(Value.q)
Define.S sMessage = "TaskbarCreated"
Define.I uTaskbarRestart = RegisterWindowMessage_(@sMessage)
;}

;{ Define Prototypes
Prototype ProcessFirst(Snapshot, Process)
Prototype ProcessNext(Snapshot, Process)
;}

;{ Global Variables
Global buffer$
Global oepath.s, BackupFile.s, MyLocation.s
Global flip1 ;Close to System Tray
Global flip2 ;Re-open OE Classic after backup completion
Global flip3 ;Open to System Tray
Global flip4 ;Window stays on top
Global windowfound
Global ProcessFirst.ProcessFirst
Global ProcessNext.ProcessNext
Global Language.Language
Global schedule.s, day.s, hour.s, minute.s, scheduledday.s
Global Dots.s
;}

;{ Window/Gadget Enumeration
Enumeration
#Window_Main
#Timer
#Panel1
#StatusBar
;Main Tab
#Text_Path
#String_Path
#Text_BackupPath
#String_BackupPath
#Button_Backup
#Button_SetBackupPath
#Button_OpenBackupLocation
#Button_RestoreBackup
;OptionsTab
#Checkbox_CloseToTray
#Checkbox_RestartOEClassic
#Checkbox_StartInTray
#Checkbox_StayOnTop
#Button_CreateDesktopIcon
#Button_OpenBackupLog
;Task Tab
#Frame_TaskSettings
#Option_Daily
#Option_Weekly
#Option_Monthly
#Option_Sun
#Option_Mon
#Option_Tue
#Option_Wed
#Option_Thu
#Option_Fri
#Option_Sat
#Option_First
#Option_Second
#Option_Third
#Option_Fourth
#Option_Last
#Option_LastDay
#SpinGadget_Hour
#TextGadget_Colon
#SpinGadget_Minute
#Text_Time
#Text_Month
#Button_CreateTask
#Button_DeleteTask
;Pop-Up Menu
#Menu_SysTray
#RestoreApp
#BackupDB
#RestoreDB
#OpenBackupFolder
#QuitApp
;Pop-Up Menu Icons
#Icon_SysTray
#Icon_RestoreApp
#Icon_BackupDB
#Icon_RestoreDB
#Icon_BackupFolder
#Icon_Quit
EndEnumeration
;}

;{ Declare Procedures
Declare BackupDatabase()
Declare FindWin(Title$)
Declare.s GetLocale()
Declare.s GetPidProcessEx(Name.s)
Declare WriteLog(filename.s, error.s)
;}

;{ Command Line Procedures
Procedure BackupDatabase()
GetLocale()
Protected GetDate.s, myid, PCName.s, MyLocation.s, MyName.s, OEInstalled.s, oepath.s
PCName=GetEnvironmentVariable("COMPUTERNAME")
MyName=GetEnvironmentVariable("USERNAME")
If FindWin("OE Classic")
  RunProgram("taskkill","/f /im oeclassic.exe","",#PB_Program_Hide|#PB_Program_Wait)
   windowfound=1
EndIf
GetDate=FormatDate("%yyyy%mm%dd %hh:%ii:%ss",Date())
OpenPreferences("oebackup.prefs")
 MyLocation=ReadPreferenceString("BkUpDir","")
ClosePreferences()
If MyLocation=""
  WriteLog("Backup",Language(Language,"LogMessages","LogNoPath"))
   HideWindow(#Window_Main,0)
    MessageRequester(Language(Language,"Messages","Error"),Language(Language,"Messages","OEPathMsg"),#MB_ICONERROR)
     windowfound=0
      End
EndIf
oepath.s=GetEnvironmentVariable("userprofile")+"\Appdata\Local\OEClassic"
 If FindString(MyLocation," ",1)
   myid=RunProgram("7z.exe","a -mmt -mx=9 -slp -y "+Chr(34)+MyLocation+Chr(34)+"OEClassicBackup_"+PCName+"_"+MyName+"_"+GetDate+".7z "+oepath,"",#PB_Program_Open|#PB_Program_Error|#PB_Program_Hide)
 Else
   myid=RunProgram("7z.exe","a -mmt -mx=9 -slp -y "+MyLocation+"OEClassicBackup_"+PCName+"_"+MyName+"_"+GetDate+".7z "+oepath,"",#PB_Program_Open|#PB_Program_Error|#PB_Program_Hide)
 EndIf
While ProgramRunning(myid)
Wend
WriteLog("Backup",Language(Language,"LogMessages","LogSuccess")+Chr(34)+MyLocation+Chr(34))
OpenPreferences("oebackup.prefs")
 If ReadPreferenceInteger("ReopenOE", 0)=1
   RunProgram("oeclassic.exe")
    windowfound=0
 EndIf
ClosePreferences()
 End
EndProcedure
;}

;{ Procedures
Procedure CheckForTask()
Protected output$, myid, taskfound
output$=""
myid=RunProgram("schtasks","/query","",#PB_Program_Open|#PB_Program_Read|#PB_Program_Hide|#PB_Ascii)
If IsProgram(myid)
 While ProgramRunning(myid)
   output$=ReadProgramString(myid)
  If FindString(output$,Language(Language,"TaskWindow","TaskName"),1,#PB_String_NoCase);"BackupOEClassic",1,#PB_String_NoCase)
    taskfound=1
     Break
  Else
    taskfound=0
  EndIf
 Wend
  CloseProgram(myid)
EndIf
ProcedureReturn taskfound
EndProcedure

Procedure.i CreateShortcut (sTargetPath.s, sLinkPath.s, sTargetArgument.s, sTargetDescription.s, sWorkingDirectory.s, iShowCommand, sIconFile.s, iIconIndexInFile)

;/***P
;*
;* DESC
;*    Create a shortcut for the specified target.
;*
;* IN
;*  sTargetPath         ; The full pathname of the target.
;*  sLinkPath           ; The full pathname of the shortcut link (the actual .lnk file).
;*  sTargetArgument     ; The argument(s) to be passed to the target.
;*  sTargetDescription  ; A description of the shortcut (it will be visible as a tooltip).
;*  sWorkingDirectory   ; The desired working directory for the target.
;*  iShowCommand        ; Start mode (SW_SHOWNORMAL, SW_SHOWMAXIMIZED, SW_SHOWMINIMIZED)
;*  sIconFile           ; The full pathname of the file containing the icon to be used (usually the target itself).
;*  iIconIndexInFile    ; The index for the icon to be retrieved form the icon file.
;*
;* RET
;*  0 OK
;*  1 FAILED
;*
;* EXAMPLE
;*  CreateShortcut ("C:\temp\program.exe", "C:\temp\program.lnk", "arg","A nice program.","c:\temp", #SW_SHOWMAXIMIZED, "C:\temp\program.exe", 0)
;* 
;* OS
;*  Windows
;***/ 

 Protected psl.IShellLinkW, ppf.IPersistFile
 Protected sBuffer.s
 Protected iRetVal = 1

 CoInitialize_(#Null)
  
 If CoCreateInstance_(?CLSID_ShellLink,0,1,?IID_IShellLink,@psl) = #S_OK
   
    ; The file TO which is linked ( = target for the Link )
    
    psl\SetPath(sTargetPath)
   
    ; Arguments for the Target
    
    psl\SetArguments(sTargetArgument)
   
    ; Working Directory
    
    psl\SetWorkingDirectory(sWorkingDirectory)
   
    ; Description ( also used as Tooltip for the Link )
    
    psl\SetDescription(sTargetDescription)
   
    ; Show command:
    ;   SW_SHOWNORMAL    = Default
    ;   SW_SHOWMAXIMIZED = Maximized
    ;   SW_SHOWMINIMIZED = Minimized
    
    psl\SetShowCmd(iShowCommand)
   
    ; Hotkey (not implemented): 
    ; The virtual key code is in the low-order byte,
    ; and the modifier flags are in the high-order byte.
    ; The modifier flags can be a combination of the following values:
    ;
    ;   HOTKEYF_ALT     = ALT key
    ;   HOTKEYF_CONTROL = CTRL key
    ;   HOTKEYF_EXT     = Extended key
    ;   HOTKEYF_SHIFT   = SHIFT key
   
    psl\SetHotkey(#Null)
   
    ; Icon for the Link:
    ; There can be more than 1 icons in an icon resource file,
    ; so you have to specify the index.
    
    psl\SetIconLocation(sIconFile, iIconIndexInFile)
   

    ; Query IShellLink for the IPersistFile interface for saving the
    ; shortcut in persistent storage.
   
    If psl\QueryInterface(?IID_IPersistFile, @ppf) = 0
    
        ; Ensure that the string is Unicode.
        sBuffer = Space(#MAX_PATH) 
      
        PokeS(@sBuffer, sLinkPath, -1, #PB_Unicode)
      
        ;Save the link by calling IPersistFile::Save.
        ppf\Save(sBuffer, #True)
      
        iRetVal = 0
      
        ppf\Release()
    EndIf
    
    psl\Release()
    
 EndIf
  
 CoUninitialize_()
  
 ProcedureReturn iRetVal
   
 DataSection
 
    CLSID_ShellLink:
    ; 00021401-0000-0000-C000-000000000046
    Data.l $00021401
    Data.w $0000,$0000
    Data.b $C0,$00,$00,$00,$00,$00,$00,$46
       
    IID_IPersistFile:
    ; 0000010b-0000-0000-C000-000000000046
    Data.l $0000010B
    Data.w $0000,$0000
    Data.b $C0,$00,$00,$00,$00,$00,$00,$46
   
CompilerIf #PB_Compiler_Unicode = 0   
    IID_IShellLink:
    ; DEFINE_SHLGUID(IID_IShellLinkA, 0x000214EEL, 0, 0);
    ; C000-000000000046
    Data.l $000214EE
    Data.w $0000,$0000
    Data.b $C0,$00,$00,$00,$00,$00,$00,$46   
CompilerElse   
    IID_IShellLink: ; {000214F9-0000-0000-C000-000000000046}
    Data.l $000214F9
    Data.w $0000, $0000
    Data.b $C0, $00, $00, $00, $00, $00, $00, $46
CompilerEndIf

 EndDataSection   
   
EndProcedure

Procedure WriteLog(filename.s, error.s)
 OpenFile(0,filename+".log",#PB_File_SharedRead|#PB_File_SharedWrite|#PB_File_Append)
  WriteStringN(0, error.s+" "+FormatDate("%mm/%dd/%yyyy"+" "+"%hh:%ii:%ss" ,Date()), #PB_Ascii)
 CloseFile(0)
EndProcedure

Procedure FindWin(Title$)
  text$ = Space(#MAX_PATH)
  Repeat
    hwnd = FindWindowEx_(0,hwnd,0,0)     
    GetWindowText_(hwnd,@Text$,#MAX_PATH)
    If FindString(Text$, Title$, 1, #PB_String_NoCase) <> 0
       findwin = hWnd
       Break
     EndIf
     ;Delay(1)
     x + 1
  Until findwin Or x > 1000
  ProcedureReturn findwin
EndProcedure

Procedure.s GetLocale()
Protected buffer$
buffer$=Space(999)
GetLocaleInfo_(#LOCALE_USER_DEFAULT,#LOCALE_SLANGUAGE,buffer$,Len(buffer$))

Select buffer$

  Case "Deutsch (Deutschland)", "German (Germany)"
    LoadLanguage(Language.Language,GetCurrentDirectory()+"\Lang\german.lng")
  Case "Dutch (Netherlands)", "Nederlands (Nederland)"
    LoadLanguage(Language.Language,GetCurrentDirectory()+"\Lang\dutch.lng")
  Default
    If buffer$ = "English (United States)"
      LoadLanguage(Language)
    Else
      MessageRequester("Warning","Language not found, defaulting to English.",#MB_ICONWARNING)
    EndIf

EndSelect

EndProcedure

Procedure.s GetPidProcessEx(Name.s)
  ;/// Return all process id as string separate by comma
  ;/// Author : jpd
  Protected ProcLib
  Protected ProcName.s
  Protected Process.PROCESSENTRY32
  Protected x
  Protected retval=#False
  Name=UCase(Name.s)
  ProcLib= OpenLibrary(#PB_Any, "Kernel32.dll") 
  If ProcLib
    CompilerIf #PB_Compiler_Unicode
      ProcessFirst           = GetFunction(ProcLib, "Process32FirstW") 
      ProcessNext            = GetFunction(ProcLib, "Process32NextW") 
    CompilerElse
      ProcessFirst           = GetFunction(ProcLib, "Process32First") 
      ProcessNext            = GetFunction(ProcLib, "Process32Next") 
    CompilerEndIf
    If  ProcessFirst And ProcessNext 
      Process\dwSize = SizeOf(PROCESSENTRY32) 
      Snapshot =CreateToolhelp32Snapshot_(#TH32CS_SNAPALL,0)
      If Snapshot 
        ProcessFound = ProcessFirst(Snapshot, Process) 
        x=1
        While ProcessFound 
          ProcName=PeekS(@Process\szExeFile)
          ProcName=GetFilePart(ProcName)
          If UCase(ProcName)=UCase(Name)
            If ProcessList.s<>"" : ProcessList+",": EndIf
            ProcessList+Str(Process\th32ProcessID)
          EndIf
          ProcessFound = ProcessNext(Snapshot, Process) 
          x=x+1  
        Wend 
      EndIf 
      CloseHandle_(Snapshot) 
    EndIf 
    CloseLibrary(ProcLib) 
  EndIf 
  ProcedureReturn ProcessList

EndProcedure

Procedure CreateBackup()
Protected GetDate.s, myid, PCName.s, closeme, MyName.s
PCName=GetEnvironmentVariable("COMPUTERNAME")
 If GetPidProcessEx("OEClassic.exe")
   closeme=MessageRequester(Language(Language,"Messages","Error"),Language(Language,"Messages","CloseError"),#PB_MessageRequester_YesNo|#MB_ICONERROR)
    windowfound=1
  If closeme=#PB_MessageRequester_Yes
    RunProgram("taskkill","/f /im oeclassic.exe","",#PB_Program_Hide)
     Delay(500)
      Goto runbackup
  Else
    WriteLog("Backup",Language(Language,"LogMessages","LogFailure"))
  EndIf
 Else
runbackup:
   GetDate=FormatDate("%yyyy%mm%dd",Date())
   MyName=GetEnvironmentVariable("USERNAME")
     StatusBarText(#StatusBar,0,Language(Language,"Messages","StatusBackup"),#PB_StatusBar_Center)
      If FindString(MyLocation," ",1)
        myid=RunProgram("7z.exe","a -mmt -mx=9 -slp -y "+Chr(34)+MyLocation+Chr(34)+"OEClassicBackup_"+PCName+"_"+MyName+"_"+GetDate+".7z "+oepath,"",#PB_Program_Open|#PB_Program_Hide)
      Else
        myid=RunProgram("7z.exe","a -mmt -mx=9 -slp -y "+MyLocation+"OEClassicBackup_"+PCName+"_"+MyName+"_"+GetDate+".7z "+oepath,"",#PB_Program_Open|#PB_Program_Hide)
      EndIf
       While ProgramRunning(myid)
        Select WaitWindowEvent(1)
          Case #PB_Event_Timer
            If ProgramRunning(myid)
              StatusBarText(#StatusBar, 0, Language(Language,"Messages","StatusBackup") + Dots,#PB_StatusBar_Center)
               Dots + "."
             If Len(Dots) >= 11
               Dots = ""
             EndIf
            Else
              StatusBarText(#StatusBar, 0, Language(Language,"MainWindow","StatusBarText"),#PB_StatusBar_Center)
            EndIf
        EndSelect
       Wend
         StatusBarText(#StatusBar,0,"Ready",#PB_StatusBar_Center)
          WriteLog("Backup",Language(Language,"LogMessages","LogSuccess")+Chr(34)+MyLocation+Chr(34))
           OpenPreferences("oebackup.prefs")
            If ReadPreferenceInteger("ReopenOE", 0)=1
              MessageRequester(Language(Language,"Messages","SuccessMsg"),Language(Language,"Messages","MsgSuccess"),#MB_ICONINFORMATION)
               RunProgram("oeclassic.exe")
                windowfound=0
            EndIf
           ClosePreferences()
 EndIf
EndProcedure

Procedure CreateBackupTask(schedule.s, day.s, scheduledday.s, hour.s, minute.s)
If GetGadgetText(#String_BackupPath)=""
  MessageRequester(Language(Language,"Messages","Error"),Language(Language,"LogMessages","LogNoPath"),#MB_ICONERROR)
   ProcedureReturn
EndIf
If schedule = "Daily"
  RunProgram("schtasks","/create /sc daily /st "+hour+":"+minute+" /tn "+Language(Language,"TaskWindow","TaskName")+" /tr "+Chr(34)+GetCurrentDirectory()+"OEClassicBackup /b"+Chr(34)+" /it /v1","",#PB_Program_Wait)
ElseIf schedule = "Monthly"
  If scheduledday.s="";1st day of every month
    RunProgram("schtasks","/create /sc monthly /st "+hour+":"+minute+" /tn "+Language(Language,"TaskWindow","TaskName")+" /tr "+Chr(34)+GetCurrentDirectory()+"OEClassicBackup /b"+Chr(34)+" /it /v1","",#PB_Program_Wait)    
  ElseIf scheduledday.s<>"LastDay";chosen day (ex: 1st sun of every month)
    RunProgram("schtasks","/create /sc monthly /d "+day+" /mo "+scheduledday+" /st "+hour+":"+minute+" /tn "+Language(Language,"TaskWindow","TaskName")+" /tr "+Chr(34)+GetCurrentDirectory()+"OEClassicBackup /b"+Chr(34)+" /it /v1","",#PB_Program_Wait)
  Else;last day of every month
    RunProgram("schtasks","/create /sc monthly /mo lastday /m * /st "+hour+":"+minute+" /tn "+Language(Language,"TaskWindow","TaskName")+" /tr "+Chr(34)+GetCurrentDirectory()+"OEClassicBackup /b"+Chr(34)+" /it /v1","",#PB_Program_Wait)
  EndIf
Else ;schedule=Weekly
  RunProgram("schtasks","/create /sc "+schedule+" /d "+day+" /st "+hour+":"+minute+" /tn "+Language(Language,"TaskWindow","TaskName")+" /tr "+Chr(34)+GetCurrentDirectory()+"OEClassicBackup /b"+Chr(34)+" /it /v1","",#PB_Program_Wait)
EndIf
 If CheckForTask()=1
   WriteLog("Backup",Language(Language,"LogMessages","TaskSucceed"))
    DisableGadget(#Option_Daily,1)
     DisableGadget(#Option_Weekly,1)
      DisableGadget(#Option_Monthly,1)
       DisableGadget(#Option_Sun,1)
        DisableGadget(#Option_Mon,1)
         DisableGadget(#Option_Tue,1)
          DisableGadget(#Option_Wed,1)
           DisableGadget(#Option_Thu,1)
            DisableGadget(#Option_Fri,1)
             DisableGadget(#Option_Sat,1)
            DisableGadget(#Option_First,1)
           DisableGadget(#Option_Second,1)
          DisableGadget(#Option_Third,1)
         DisableGadget(#Option_Fourth,1)
        DisableGadget(#Option_Last,1)
       DisableGadget(#Option_LastDay,1)
      DisableGadget(#SpinGadget_Hour,1)
     DisableGadget(#TextGadget_Colon,1)
    DisableGadget(#SpinGadget_Minute,1)
   DisableGadget(#Button_CreateTask,1)
  DisableGadget(#Text_Time,1)
  HideGadget(#Text_Month,1)
 HideGadget(#Button_DeleteTask,0)
 Else
   WriteLog("Backup",Language(Language,"LogMessages","TaskFail"))
    MessageRequester(language(Language,"Messages","Error"),language(Language,"LogMessages","TaskFail"))
 EndIf
EndProcedure

Procedure OpenBackupLocation()
 If MyLocation<>""
  If FindString(MyLocation," ",1)
    RunProgram("explorer",Chr(34)+MyLocation+Chr(34),"")
  Else
    RunProgram("explorer",MyLocation,"")
  EndIf
 EndIf
EndProcedure

Procedure RestoreBackup()
Protected myid, restorepath.s
 If GetPidProcessEx("OEClassic.exe")
   MessageRequester(Language(Language,"Messages","Error"),Language(Language,"Messages","ErrNotClosed"),#MB_ICONERROR)
    WriteLog("Backup",Language(Language,"LogMessages","LogFailure2"))
 Else
   myrestore=MessageRequester(Language(Language,"RestoreMsgs","Restore"),Language(Language,"RestoreMsgs","RestoreQuest"),#PB_MessageRequester_YesNo|#MB_ICONQUESTION)
  If myrestore=#PB_MessageRequester_Yes
   If MyLocation<>""
     startpath.s=MyLocation
   Else
     startpath.s=""
   EndIf
     BackupFile=OpenFileRequester(language(Language,"RestoreMsgs","RestoreBkup"),startpath,"7-Zip (*.7z)|*.7z",0)
    If BackupFile<>""
      yesrestore=MessageRequester(Language(Language,"RestoreMsgs","Warning"),Language(Language,"RestoreMsgs","Warning1"),#PB_MessageRequester_YesNo|#MB_ICONEXCLAMATION)
     If yesrestore=#PB_MessageRequester_Yes
       restorepath.s=GetEnvironmentVariable("USERPROFILE")
        restorepath+"\AppData\Local\"
         DeleteDirectory(oepath,"*.*",#PB_FileSystem_Recursive|#PB_FileSystem_Force)
          StatusBarText(#StatusBar,0,Language(Language,"RestoreMsgs","StatusText1"),#PB_StatusBar_Center)
      If FindString(BackupFile," ",1)
        myid=RunProgram("7z.exe","x -o"+restorepath+" "+Chr(34)+BackupFile+Chr(34),"",#PB_Program_Open|#PB_Program_Hide|#PB_Program_Wait)
      Else
        myid=RunProgram("7z.exe","x -o"+restorepath+" "+BackupFile,"",#PB_Program_Open|#PB_Program_Hide|#PB_Program_Wait)
      EndIf
       While WaitWindowEvent(1):Wend
        While ProgramRunning(myid)
          Debug "Restore Running"
        Wend
        StatusBarText(#Statusbar,0,Language(Language,"MainWindow","StatusBarText"),#PB_StatusBar_Center)
       WriteLog("Backup",Language(Language,"LogMessages","LogSuccess2"))
      MessageRequester(Language(Language,"Messages","SuccessMsg"),Language(Language,"RestoreMsgs","Completed"),#MB_ICONINFORMATION)
     Else
       MessageRequester(Language(Language,"RestoreMsgs","Cancelled"),Language(Language,"RestoreMsgs","Cancelled2"),#MB_ICONINFORMATION)
     EndIf
    EndIf
  Else
    MessageRequester(Language(Language,"RestoreMsgs","Cancelled"),Language(Language,"RestoreMsgs","Cancelled2"),#MB_ICONINFORMATION)
  EndIf
 EndIf
EndProcedure

Procedure WinCallback(hWnd, uMsg, WParam, LParam) 
  
  Shared uTaskbarRestart
  
  If uMsg = uTaskbarRestart
    ; You need to alter the parameters to provide the right window number (the first zero, second parameter, in this line).
    PostEvent(#TaskbarCreated, #Window_Main, 0) 
  EndIf
  
  ProcedureReturn #PB_ProcessPureBasicEvents 
  
EndProcedure
;}

;{ Create Preference File
If OpenPreferences("oebackup.prefs")=0
  CreatePreferences("oebackup.prefs")
   OpenPreferences("oebackup.prefs")
    WritePreferenceString("BkUpDir","")
    WritePreferenceInteger("CloseToTray",0)
    WritePreferenceInteger("OpenToTray",0)
    WritePreferenceInteger("ReopenOE",0)
    WritePreferenceInteger("StayOnTop",0)
   ClosePreferences()
EndIf
;}

;{ Extract 7-Zip Files from Executable
If FileSize("7z.exe")=-1
CreateFile(0,"7z.exe")
OpenFile(0,"7z.exe")
 WriteData(0,?Start7zipexe,?End7zipexe-?Start7zipexe)
CloseFile(0)
EndIf
If FileSize("7z.dll")=-1
CreateFile(1,"7z.dll")
OpenFile(1,"7z.dll")
 WriteData(1,?Start7zipdll,?End7zipdll-?Start7zipdll)
CloseFile(1)
EndIf
;}

;{ Command Line Options
If ProgramParameter()<>""
Select Left(ProgramParameter(0),2)
  Case "/b" ;Backup Database
    BackupDatabase()
EndSelect
EndIf
;}

;{ Read Preferences
OpenPreferences("oebackup.prefs")
MyLocation=ReadPreferenceString("BkUpDir","")
CloseToTray=ReadPreferenceInteger("CloseToTray",0)
OpenToTray=ReadPreferenceInteger("OpenToTray",0)
ReopenMyOE=ReadPreferenceInteger("ReopenOE",0)
StayOnTop=ReadPreferenceInteger("StayOnTop",0)
ClosePreferences()
;}

;{ Check for Lang folder and set language
If FileSize(GetCurrentDirectory()+"Lang")=-1
 MessageRequester("Error","Language folder missing, defaulting to English.",#MB_ICONWARNING)
EndIf
GetLocale()
;}

;{ Create mutex
MutexID=CreateMutex_(0,1,"OE Classic Backup")
MutexError=GetLastError_()
If MutexID=0 Or MutexError<>0
  MessageRequester(Language(Language,"Messages","Error"),Language(Language,"Misc","AppRunning"),#MB_ICONWARNING)
  End
EndIf
;}

;{ Create Window and Gadgets
If FileSize(GetEnvironmentVariable("userprofile")+"\AppData\Local\OEClassic")<>-1
  oepath.s=GetEnvironmentVariable("userprofile")+"\AppData\Local\OEClassic"
Else
  MessageRequester(Language(Language,"Messages","Error"),Language(Language,"Messages","OEPathMsg"),#MB_ICONERROR)
   End
EndIf
If ProgramParameter()=""
 If OpenToTray=1
   startme=#PB_Window_Invisible|#PB_Window_ScreenCentered
 Else
   startme=#PB_Window_ScreenCentered
 EndIf
  If OpenWindow(#Window_Main,0,0,500,200,Language(Language,"MainWindow","WinTitle"),startme|#PB_Window_MinimizeGadget)
   PanelGadget(#Panel1,5,5,492,170)
    AddGadgetItem(#Panel1,0,Language(Language,"MainWindow","FirstTab"))
     TextGadget(#Text_Path,40,7,100,20,Language(Language,"MainWindow","OEPath"))
      StringGadget(#String_Path,40,25,400,20,oepath,#PB_String_ReadOnly)
       HyperLinkGadget(#Text_BackupPath,40,47,290,20,Language(Language,"MainWindow","BackupLocation"),#Blue)
       GadgetToolTip(#Text_BackupPath,Language(Language,"MainWindow","BackupClearTip"))
        StringGadget(#String_BackupPath,40,67,400,20,MyLocation)
         ButtonGadget(#Button_Backup,0,105,150,30,Language(Language,"MainWindow","ButtonCreateBackup"))
         ButtonGadget(#Button_SetBackupPath,167,105,150,30,Language(Language,"MainWindow","ButtonSetLocation"))
         ButtonGadget(#Button_OpenBackupLocation,167,105,150,30,Language(Language,"MainWindow","ButtonOpenLocation"))
         ButtonGadget(#Button_RestoreBackup,334,105,150,30,Language(Language,"MainWindow","ButtonRestore"))
          CreateStatusBar(#StatusBar,WindowID(#Window_Main))
           AddStatusBarField(500)
           StatusBarText(#StatusBar,0,Language(Language,"MainWindow","StatusBarText"),#PB_StatusBar_Center)
    CloseGadgetList()
    OpenGadgetList(#Panel1)
     AddGadgetItem(#Panel1,1,Language(Language,"OptionsWindow","SecondTab"));Options
      CheckBoxGadget(#Checkbox_CloseToTray,20,10,400,20,Language(Language,"OptionsWindow","CheckClose")); Flip1
      CheckBoxGadget(#Checkbox_RestartOEClassic,20,30,400,20,Language(Language,"OptionsWindow","CheckRestartOE"));Flip2
      CheckBoxGadget(#Checkbox_StartInTray,20,50,400,20,Language(Language,"OptionsWindow","CheckStartInTray"));Flip3
      CheckBoxGadget(#Checkbox_StayOnTop,20,70,400,20,Language(Language,"OptionsWindow","CheckStayOnTop"));Flip4
       ButtonGadget(#Button_CreateDesktopIcon,0,105,155,30,Language(Language,"OptionsWindow","ButtonDesktopShortcut"))
       ButtonGadget(#Button_OpenBackupLog,328,105,155,30,Language(Language,"OptionsWindow","ButtonOpenLog"))
    CloseGadgetList()
    OpenGadgetList(#Panel1)
     AddGadgetItem(#Panel1,2,Language(Language,"TaskWindow","ThirdTab"));Backup Task
       FrameGadget(#Frame_TaskSettings,5,5,475,95,Language(Language,"TaskWindow","TaskFrame"))
        OptionGadget(#Option_Daily,20,25,100,20,Language(Language,"TaskWindow","OptionDaily"))
        OptionGadget(#Option_Weekly,20,50,100,20,Language(Language,"TaskWindow","OptionWeekly"))
        OptionGadget(#Option_Monthly,20,75,100,20,Language(Language,"TaskWindow","OptionMonthly"))
         SpinGadget(#SpinGadget_Hour,130,70,50,23,0,23,#PB_Spin_Numeric)
          SetGadgetText(#SpinGadget_Hour,"0")
           TextGadget(#TextGadget_Colon,190,70,5,20,":")
         SpinGadget(#SpinGadget_Minute,200,70,50,23,0,59,#PB_Spin_Numeric)
          SetGadgetText(#SpinGadget_Minute,"0")
           TextGadget(#Text_Time,255,75,200,20,Language(Language,"TaskWindow","ValidTime"))
            TextGadget(#Text_Month,20,100,250,30,Language(Language,"TaskWindow","TextMonth"))
             HideGadget(#Text_Month,1)
        OptionGadget(#Option_Sun,130,25,50,20,Language(Language,"TaskWindow","OptionSunday"))
        OptionGadget(#Option_Mon,180,25,50,20,Language(Language,"TaskWindow","OptionMonday"))
        OptionGadget(#Option_Tue,230,25,50,20,Language(Language,"TaskWindow","OptionTuesday"))
        OptionGadget(#Option_Wed,280,25,50,20,Language(Language,"TaskWindow","OptionWednesday"))
        OptionGadget(#Option_Thu,330,25,50,20,Language(Language,"TaskWindow","OptionThursday"))
        OptionGadget(#Option_Fri,380,25,50,20,Language(Language,"TaskWindow","OptionFriday"))
        OptionGadget(#Option_Sat,430,25,45,20,Language(Language,"TaskWindow","OptionSaturday"))
         ButtonGadget(#Button_CreateTask,328,105,155,30,Language(Language,"TaskWindow","ButtonCreateTask"))
         ButtonGadget(#Button_DeleteTask,0,105,150,30,Language(Language,"TaskWindow","ButtonDeleteTask"))
          HideGadget(#Button_DeleteTask,1)
        OptionGadget(#Option_First,130,50,50,20,Language(Language,"TaskWindow","OptionFirst"))
        OptionGadget(#Option_Second,180,50,50,20,Language(Language,"TaskWindow","OptionSecond"))
        OptionGadget(#Option_Third,230,50,50,20,Language(Language,"TaskWindow","OptionThird"))
        OptionGadget(#Option_Fourth,280,50,50,20,Language(Language,"TaskWindow","OptionFourth"))
        OptionGadget(#Option_Last,330,50,50,20,Language(Language,"TaskWindow","OptionLast"))
        OptionGadget(#Option_LastDay,380,50,80,20,Language(Language,"TaskWindow","OptionLastDay"))
    CloseGadgetList()
       DisableGadget(#Option_Sun,1)
        DisableGadget(#Option_Mon,1)
         DisableGadget(#Option_Tue,1)
          DisableGadget(#Option_Wed,1)
           DisableGadget(#Option_Thu,1)
            DisableGadget(#Option_Fri,1)
             DisableGadget(#Option_Sat,1)
            DisableGadget(#Option_First,1)
           DisableGadget(#Option_Second,1)
          DisableGadget(#Option_Third,1)
         DisableGadget(#Option_Fourth,1)
        DisableGadget(#Option_Last,1)
       DisableGadget(#Option_LastDay,1)
  If CreatePopupImageMenu(#Menu_SysTray);System Tray Menu
    MenuItem(#RestoreApp,Language(Language,"PopUpMenu","RestoreWin"),CatchImage(#Icon_RestoreApp,?Icon_RestoreApp))
    MenuBar()
     MenuItem(#BackupDB,Language(Language,"PopUpMenu","Backup"),CatchImage(#Icon_BackupDB,?Icon_BackupDB))
      MenuItem(#RestoreDB,Language(Language,"PopUpMenu","Restore"),CatchImage(#Icon_RestoreDB,?Icon_RestoreDB))
       MenuItem(#OpenBackupFolder,Language(Language,"PopUpMenu","OpenFolder"),CatchImage(#Icon_BackupFolder,?Icon_BackupFolder))
    MenuBar()
        MenuItem(#QuitApp,Language(Language,"PopUpMenu","QuitApp"),CatchImage(#Icon_Quit,?Icon_Quit))
  EndIf
;{ Read Preferences/Disable TaskTab Gadgets
    If CloseToTray=1
      SetGadgetState(#Checkbox_CloseToTray,#PB_Checkbox_Checked)
       flip1=1
    EndIf
    If ReopenMyOe=1
      SetGadgetState(#Checkbox_RestartOEClassic,#PB_Checkbox_Checked)
       flip2=1
    EndIf
    If OpenToTray=1
      SetGadgetState(#Checkbox_StartInTray,#PB_Checkbox_Checked)
       systrayicon.l = CatchImage(#Icon_SysTray, ?Icon_Systray)
        AddSysTrayIcon(0,WindowID(#Window_Main),systrayicon)
         SysTrayIconToolTip(0,Language(Language,"Messages","SysTrayTooltip"))
          ShowWindow_(WindowID(#Window_Main),#SW_HIDE)
           flip3=1
    EndIf
    If StayOnTop=1
      SetGadgetState(#Checkbox_StayOnTop,#PB_Checkbox_Checked)
       StickyWindow(#Window_Main,1)
        flip4=1
    EndIf
         If CheckForTask()=1
             DisableGadget(#Option_Daily,1)
              DisableGadget(#Option_Weekly,1)
               DisableGadget(#Option_Monthly,1)
                DisableGadget(#Option_Sun,1)
                 DisableGadget(#Option_Mon,1)
                  DisableGadget(#Option_Tue,1)
                   DisableGadget(#Option_Wed,1)
                    DisableGadget(#Option_Thu,1)
                     DisableGadget(#Option_Fri,1)
                      DisableGadget(#Option_Sat,1)
                       DisableGadget(#Option_First,1)
                      DisableGadget(#Option_Second,1)
                     DisableGadget(#Option_Third,1)
                    DisableGadget(#Option_Fourth,1)
                   DisableGadget(#Option_Last,1)
                  DisableGadget(#Option_LastDay,1)
                 DisableGadget(#SpinGadget_Hour,1)
                DisableGadget(#TextGadget_Colon,1)
               DisableGadget(#SpinGadget_Minute,1)
              DisableGadget(#Text_Time,1)
             DisableGadget(#Button_CreateTask,1)
            
            HideGadget(#Button_DeleteTask,0)
         EndIf
          If FileSize(GetEnvironmentVariable("userprofile")+"\Desktop\OE Classic Backup.lnk")<>-1
            DisableGadget(#Button_CreateDesktopIcon,1)
          EndIf
;}
         SetWindowCallback(@WinCallback()); Activate the callback
  EndIf
AddWindowTimer(#Window_Main,#Timer,1000)
EndIf
DisableGadget(#Button_CreateTask,1)
;}

;{ Main Loop
Repeat

;{ Disable/Hide Gadgets
If MyLocation<>""
  HideGadget(#Button_SetBackupPath,1)
   HideGadget(#Button_OpenBackupLocation,0)
    DisableGadget(#Button_Backup,0)
     DisableGadget(#Button_OpenBackupLocation,0)
      DisableGadget(#Button_RestoreBackup,0)
       DisableMenuItem(#Menu_SysTray,#BackupDB,0)
        DisableMenuItem(#Menu_SysTray,#RestoreDB,0)
         DisableMenuItem(#Menu_SysTray,#OpenBackupFolder,0)

Else
  HideGadget(#Button_SetBackupPath,0)
   HideGadget(#Button_OpenBackupLocation,1)
    DisableGadget(#Button_Backup,1)
     DisableGadget(#Button_OpenBackupLocation,1)
      DisableGadget(#Button_RestoreBackup,1)
       DisableMenuItem(#Menu_SysTray,#BackupDB,1)
        DisableMenuItem(#Menu_SysTray,#RestoreDB,1)
         DisableMenuItem(#Menu_SysTray,#OpenBackupFolder,1)
EndIf
;}

event=WaitWindowEvent(1)

Select event

  Case #PB_Event_Gadget
    eventgadget=EventGadget()
     eventtype=EventType()
      eventwindow=EventWindow()

    Select eventgadget
;{ Button gadgets
      Case #Button_Backup
        CreateBackup()

      Case #Button_CreateDesktopIcon
        CreateShortcut(GetCurrentDirectory()+"oeclassicbackup.exe",GetEnvironmentVariable("USERPROFILE")+"\Desktop\OE Classic Backup.lnk","",Language(Language,"Misc","DesktopIconText"),GetCurrentDirectory(),#SW_SHOWNORMAL,GetCurrentDirectory()+"oeclassicbackup.exe",0)
         If FileSize(GetEnvironmentVariable("userprofile")+"\Desktop\OE Classic Backup.lnk")<>-1
           DisableGadget(#Button_CreateDesktopIcon,1)
         EndIf

      Case #Button_OpenBackupLocation
        OpenBackupLocation()

      Case #Button_OpenBackupLog
        If FileSize("backup.log")<>-1
          RunProgram("backup.log")
        Else
          MessageRequester(Language(Language,"Messages","Error"),Language(Language,"Messages","BackupErrorRead"),#MB_ICONWARNING)
        EndIf

      Case #Button_RestoreBackup
        RestoreBackup()

      Case #Button_SetBackupPath
        MyLocation=PathRequester(Language(Language,"Misc","MsgBackupLocation"),GetEnvironmentVariable("homedrive"))
         If MyLocation<>""
           SetGadgetText(#String_BackupPath,MyLocation)
            OpenPreferences("oebackup.prefs")
             WritePreferenceString("BkUpDir",GetGadgetText(#String_BackupPath))
            ClosePreferences()
         EndIf

       Case #Button_CreateTask
         starthour.s=GetGadgetText(#SpinGadget_Hour)
          If Len(starthour)<2
            thehour.s=Left(starthour,1)
             hour.s=InsertString(thehour,"0",1)
          Else
            hour=starthour
          EndIf
          startminute.s=GetGadgetText(#SpinGadget_Minute)
          If Len(startminute)<2
            minute.s=InsertString(startminute,"0",1)
          Else
            minute=startminute
          EndIf
         CreateBackupTask(schedule.s, day.s, scheduledday.s, hour.s, minute.s)

       Case #Button_DeleteTask
         RunProgram("schtasks","/delete /tn "+Language(Language,"TaskWindow","TaskName")+" /f","",#PB_Program_Wait)
          If CheckForTask()=0
            DisableGadget(#Option_Daily,0)
             DisableGadget(#Option_Weekly,0)
              DisableGadget(#Option_Monthly,0)
              DisableGadget(#Button_CreateTask,0)
             HideGadget(#Button_DeleteTask,1)
            SetActiveGadget(#Option_Daily)
          EndIf
;}
;{ Checkbox Gadgets
      Case #Checkbox_CloseToTray; Flip1
        flip1=1-flip1
        If flip1=1
          OpenPreferences("oebackup.prefs")
           WritePreferenceInteger("CloseToTray",1)
          ClosePreferences()
        Else
          OpenPreferences("oebackup.prefs")
           WritePreferenceInteger("CloseToTray",0)
          ClosePreferences()
        EndIf

      Case #Checkbox_RestartOEClassic; Flip2
        flip2=1-flip2
        If flip2=1
          OpenPreferences("oebackup.prefs")
           WritePreferenceInteger("ReopenOE",1)
          ClosePreferences()
        Else
          OpenPreferences("oebackup.prefs")
           WritePreferenceInteger("ReopenOE",0)
          ClosePreferences()
        EndIf

      Case #Checkbox_StartInTray; Flip3
        flip3=1-flip3
        If flip3=1
          OpenPreferences("oebackup.prefs")
           WritePreferenceInteger("OpenToTray",1)
          ClosePreferences()
        Else
          OpenPreferences("oebackup.prefs")
           WritePreferenceInteger("OpenToTray",0)
          ClosePreferences()
        EndIf

      Case #Checkbox_StayOnTop; Flip4
        flip4=1-flip4
        If flip4=1
          OpenPreferences("oebackup.prefs")
           WritePreferenceInteger("StayOnTop",1)
          ClosePreferences()
           StickyWindow(#Window_Main,1)
        Else
          OpenPreferences("oebackup.prefs")
           WritePreferenceInteger("StayOnTop",0)
          ClosePreferences()
           StickyWindow(#Window_Main,0)
        EndIf
;}
;{ Option Gadgets
      Case #Option_Sun
        day="Sun"
      Case #Option_Mon
        day="Mon"
      Case #Option_Tue
        day="Tue"
      Case #Option_Wed
        day="Wed"
      Case #Option_Thu
        day="Thu"
      Case #Option_Fri
        day="Fri"
      Case #Option_Sat
        day="Sat"
      Case #Option_First
        scheduledday="First"
      Case #Option_Second
        scheduledday="Second"
      Case #Option_Third
        scheduledday="Third"
      Case #Option_Fourth
        scheduledday="Fourth"
      Case #Option_Last
        scheduledday="Last"
      Case #Option_LastDay
        scheduledday="LastDay"
      Case #Option_Daily
        schedule="Daily"
         DisableGadget(#Option_Sun,1)
          DisableGadget(#Option_Mon,1)
           DisableGadget(#Option_Tue,1)
            DisableGadget(#Option_Wed,1)
             DisableGadget(#Option_Thu,1)
              DisableGadget(#Option_Fri,1)
               DisableGadget(#Option_Sat,1)
                DisableGadget(#Option_First,1)
                 DisableGadget(#Option_Second,1)
                  DisableGadget(#Option_Third,1)
                   DisableGadget(#Option_Fourth,1)
                    DisableGadget(#Option_Last,1)
                     DisableGadget(#Option_LastDay,1)
                      DisableGadget(#Button_CreateTask,0)
                       HideGadget(#Text_Month,1)

      Case #Option_Weekly
        schedule="Weekly"
         DisableGadget(#Option_Sun,0)
          DisableGadget(#Option_Mon,0)
           DisableGadget(#Option_Tue,0)
            DisableGadget(#Option_Wed,0)
             DisableGadget(#Option_Thu,0)
              DisableGadget(#Option_Fri,0)
               DisableGadget(#Option_Sat,0)
                DisableGadget(#Option_First,1)
                 DisableGadget(#Option_Second,1)
                  DisableGadget(#Option_Third,1)
                   DisableGadget(#Option_Fourth,1)
                    DisableGadget(#Option_Last,1)
                     DisableGadget(#Option_LastDay,1)
                      DisableGadget(#Button_CreateTask,0)
                       HideGadget(#Text_Month,1)

      Case #Option_Monthly
        schedule="Monthly"
         DisableGadget(#Option_Sun,0)
          DisableGadget(#Option_Mon,0)
           DisableGadget(#Option_Tue,0)
            DisableGadget(#Option_Wed,0)
             DisableGadget(#Option_Thu,0)
              DisableGadget(#Option_Fri,0)
               DisableGadget(#Option_Sat,0)
                DisableGadget(#Option_First,0)
                 DisableGadget(#Option_Second,0)
                  DisableGadget(#Option_Third,0)
                   DisableGadget(#Option_Fourth,0)
                    DisableGadget(#Option_Last,0)
                     DisableGadget(#Option_LastDay,0)
                      DisableGadget(#Button_CreateTask,0)
                       HideGadget(#Text_Month,0)

;}
;{ Text Gadgets
      Case #Text_BackupPath
        SetGadgetText(#String_BackupPath,"")
         OpenPreferences("oebackup.prefs")
          WritePreferenceString("BkUpDir","")
         ClosePreferences()
          MyLocation=""
;}

    EndSelect
;{ Menu Events
    Case #PB_Event_Menu

      Select EventMenu()

        Case #RestoreApp
          ShowWindow_(WindowID(#Window_Main),#SW_RESTORE)
           RemoveSysTrayIcon(0)

        Case #BackupDB
            CreateBackup()

        Case #RestoreDB
          RestoreBackup()

        Case #OpenBackupFolder
          OpenBackupLocation()

        Case #QuitApp
          RemoveSysTrayIcon(0)
           End

      EndSelect
;}
;{ System Tray Events
    Case #PB_Event_SysTray
      DisplayPopupMenu(#Menu_SysTray, WindowID(#Window_Main))
;}
;{ Close Window Events

      Case #PB_Event_CloseWindow
        If GetGadgetState(#Checkbox_CloseToTray)=#PB_Checkbox_Checked
          systrayicon.l = CatchImage(#Icon_SysTray, ?Icon_Systray)
           AddSysTrayIcon(0,WindowID(#Window_Main),systrayicon)
            SysTrayIconToolTip(0,Language(Language,"Messages","SysTrayTooltip"))
             ShowWindow_(WindowID(#Window_Main),#SW_HIDE)
        Else
          End
        EndIf
;}
;{ Taskbar Re-created
      Case #TaskbarCreated
        RemoveSysTrayIcon(0)
         systrayicon.l=CatchImage(#Icon_SysTray,?Icon_SysTray)
          AddSysTrayIcon(0,WindowID(#Window_Main),systrayicon)
           Debug "#TaskbarCreated"
;}

EndSelect

ForEver
ResetLanguage(Language)
;}

;{ Embed Files
DataSection

Icon_RestoreApp:
IncludeBinary ".\gfx\restorewnd.ico"

Icon_BackupDB:
IncludeBinary ".\gfx\backupdb.ico"

Icon_RestoreDB:
IncludeBinary ".\gfx\restoredb.ico"

Icon_BackupFolder:
IncludeBinary ".\gfx\folder.ico"

Icon_Quit:
IncludeBinary ".\gfx\quit.ico"

Icon_Systray:
IncludeBinary ".\gfx\icon3.ico"

Start7zipexe:
IncludeBinary ".\includes\7z.exe"
End7zipexe:

Start7zipdll:
IncludeBinary ".\includes\7z.dll"
End7zipdll:

EndDataSection
;}

;{ Default Language
DataSection

  Language:

  ;==================================================================================================================
  Data.s "_GROUP_",                                                                                      "MainWindow"
  ;==================================================================================================================
    Data.s "WinTitle",                                                                            "OE Classic Backup"
    Data.s "FirstTab",                                                                                         "Main"
    Data.s "OEPath",                                                                               "OE Classic Path:"
    Data.s "BackupLocation",                                                                           "BackUp Path:"
    Data.s "BackupClearTip",                                                             "Click to clear backup path"
    Data.s "ButtonCreateBackup",                                                                             "Backup"
    Data.s "ButtonSetLocation",                                                                     "Set Backup Path"
    Data.s "ButtonOpenLocation",                                                                   "Open Backup Path"
    Data.s "ButtonRestore",                                                                                 "Restore"
    Data.s "StatusBarText",                                                                                   "Ready"

  ;==================================================================================================================
  Data.s "_GROUP_",                                                                                   "OptionsWindow"
  ;==================================================================================================================
    Data.s "SecondTab",                                                                                     "Options"
    Data.s "CheckClose",                                                                              "Close to Tray"
    Data.s "CheckRestartOE",                                                        "Restart OE Classic after backup"
    Data.s "CheckStartInTray",                                                                 "Start in System Tray"
    Data.s "CheckStayOnTop",                                                                            "Stay on Top"
    Data.s "ButtonDesktopShortcut",                                                         "Create Desktop Shortcut"
    Data.s "ButtonOpenLog",                                                                                "View Log"

  ;==================================================================================================================
  Data.s "_GROUP_",                                                                                      "TaskWindow"
  ;==================================================================================================================
    Data.s "ThirdTab",                                                                                  "Backup Task"
    Data.s "ButtonCreateTask",                                                                          "Create Task"
    Data.s "ButtonDeleteTask",                                                                          "Remove Task"
    Data.s "TaskName",                                                                              "BackupOEClassic"
    Data.s "TaskFrame",                                                                               "Task Settings"
    Data.s "OptionDaily",                                                                                     "Daily"
    Data.s "OptionWeekly",                                                                                   "Weekly"
    Data.s "OptionMonthly",                                                                                 "Monthly"
    Data.s "TextMonth",                                 "Select time, but leave all else blank for 1st of each month"
    Data.s "ValidTime",                                                         "(Valid time is from 00:00 to 23:59)"
    Data.s "OptionSunday",                                                                                      "Sun"
    Data.s "OptionMonday",                                                                                      "Mon"
    Data.s "OptionTuesday",                                                                                     "Tue"
    Data.s "OptionWednesday",                                                                                   "Wed"
    Data.s "OptionThursday",                                                                                    "Thu"
    Data.s "OptionFriday",                                                                                      "Fri"
    Data.s "OptionSaturday",                                                                                    "Sat"
    Data.s "OptionFirst",                                                                                       "1st"
    Data.s "OptionSecond",                                                                                      "2nd"
    Data.s "OptionThird",                                                                                       "3rd"
    Data.s "OptionFourth",                                                                                      "4th"
    Data.s "OptionLast",                                                                                       "Last"
    Data.s "OptionLastDay",                                                                                "Last Day"
  ;==================================================================================================================
  Data.s "_GROUP_",                                                                                       "PopUpMenu"
  ;==================================================================================================================
    Data.s "RestoreWin",                                                                             "Restore Window"
    Data.s "Backup",                                                                                "Backup Database"
    Data.s "Restore",                                                                              "Restore Database"
    Data.s "OpenFolder",                                                                         "Open Backup Folder"
    Data.s "QuitApp",                                                                                          "Quit"

  ;==================================================================================================================
  Data.s "_GROUP_",                                                                                        "Messages"
  ;==================================================================================================================
    Data.s "BackupErrorRead",                   "Cannot open backup log. File has been deleted or no backup ran yet."
    Data.s "CloseError",                                 "OE Classic must be closed before backing up. Close it now?"
    Data.s "Error",                                                                                           "Error"
    Data.s "ErrNotClosed",                                         "OE Classic must be closed before restoring data."
    Data.s "MsgSuccess",                                                   "Backup Completed, Re-starting OE Classic"
    Data.s "OEPathMsg",                                                                "OE Classic path not found!!!"
    Data.s "StatusBackup",                                                                          "Creating backup"
    Data.s "SuccessMsg",                                                                                    "Success"
    Data.s "SysTrayTooltip",                                                  "OE Classic Backup - Click for Options"

  ;==================================================================================================================
  Data.s "_GROUP_",                                                                                            "Misc"
  ;==================================================================================================================
    Data.s "AppRunning",                          "Application is already running!!! Please check the system tray!!!"
    Data.s "DesktopIconText",                                                                     "OE Classic Backup"
    Data.s "MsgBackupLocation",                                                             "Choose Backup Location:"
    Data.s "TaskName",                                                                              "BackupOEClassic"

  ;==================================================================================================================
  Data.s "_GROUP_",                                                                                     "RestoreMsgs"
  ;==================================================================================================================
    Data.s "RestoreQuest",                                          "Are you sure you wish to restore from a backup?"
    Data.s "Restore",                                                                                       "Restore"
    Data.s "RestoreBkup",                                                                 "Select backup to restore:"
    Data.s "Warning",                                                                                       "Warning"
    Data.s "Warning1",         "This will overwrite your current OE Classic Data. Are you sure you wish to continue?"
    Data.s "StatusText1",                                                            "Please Wait.....restoring data"
    Data.s "Completed",                                                                           "Restore Completed"
    Data.s "Cancelled",                                                                                   "Cancelled"
    Data.s "Cancelled2",                                                   "Your restore request has been cancelled."

  ;==================================================================================================================
  Data.s "_GROUP_",                                                                                     "LogMessages"
  ;==================================================================================================================
    Data.s "LogNoPath",                                                                     "No backup path defined."
    Data.s "LogFailure",                                                  "Backup failed, OE Classic was not closed."
    Data.s "LogFailure2",                                                "Restore failed, OE Classic was not closed."
    Data.s "LogSuccess",                                                     "Database backup successfully saved to "
    Data.s "LogSuccess2",                                                                       "Restore Successful."
    Data.s "TaskFail",                                                                  "Scheduled task not created."
    Data.s "TaskSucceed",                                                      "Scheduled task successfully created."

  ;==================================================================================================================
  Data.s "_END_",                                                                                                  ""
  ;==================================================================================================================
  
EndDataSection
;}
; IDE Options = PureBasic 6.04 LTS (Windows - x64)
; CursorPosition = 138
; FirstLine = 12
; Folding = AngAAAA-
; EnableThread
; EnableXP
; DPIAware
; UseIcon = gfx\icon3.ico
; Executable = C:\Temp\OEClassicBackup.exe
; Debugger = Standalone