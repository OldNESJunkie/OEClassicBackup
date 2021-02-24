﻿;*********************************
;**  OE Classic Database Backup **
;**              by             **
;**   Daniel Ford 02/12/2021    **
;*********************************

;{ Define Prototypes
Prototype ProcessFirst(Snapshot, Process)
Prototype ProcessNext(Snapshot, Process)
;}

;{ Global Variables
Global oepath.s, BackupFile.s, MyLocation.s, flip1, flip2, flip3, flip4
Global ProcessFirst.ProcessFirst
Global ProcessNext.ProcessNext
;}

;{ Window/Gadget Enumeration
Enumeration
#Window_Main
#Panel1
#Text_Path
#String_Path
#Text_BackupPath
#String_BackupPath
#Button_Backup
#Button_CreateDesktopIcon
#Button_OpenBackupLocation
#Button_RestoreBackup
#Button_SetBackupPath
#Button_SetTask
#Checkbox_RememberBackupLocation
#Checkbox_CloseToTray
#Checkbox_StartInTray
#Checkbox_StartWithLogon
#StatusBar
#Menu_SysTray
#RestoreApp
#BackupDB
#RestoreDB
#OpenBackupFolder
#QuitApp
#Icon_SysTray
#Icon_RestoreApp
#Icon_BackupDB
#Icon_RestoreDB
#Icon_BackupFolder
#Icon_Quit
EndEnumeration
;}

;{ Declare Procedures
Declare.s GetPidProcessEx(Name.s)
Declare WriteLog(filename.s, error.s)
;}

;{ Command Line Procedures
Procedure BackupDatabase()
Protected GetDate.s, myid
RunProgram("taskkill","/f /im oeclassic.exe","",#PB_Program_Hide|#PB_Program_Wait)
GetDate=FormatDate("%yyyy%mm%dd",Date())
OpenPreferences("oebackup.prefs")
 MyLocation=ReadPreferenceString("BkUpDir","")
ClosePreferences()
If MyLocation=""
  WriteLog("Backup","No backup path defined")
   HideWindow(#Window_Main,0)
    MessageRequester("Error","No backup path defined",#MB_ICONERROR)
     ProcedureReturn
EndIf
oepath.s=GetEnvironmentVariable("userprofile")+"\Appdata\Local\OEClassic"
 If FindString(MyLocation," ",1)
   myid=RunProgram("7z.exe","a -mx=9 "+Chr(34)+MyLocation+Chr(34)+"OEClassicBackup_"+GetDate+".7z "+oepath,"",#PB_Program_Open|#PB_Program_Hide)
 Else
   myid=RunProgram("7z.exe","a -mx=9 "+MyLocation+"OEClassicBackup_"+GetDate+".7z "+oepath,"",#PB_Program_Open|#PB_Program_Hide)
 EndIf
While ProgramRunning(myid)
Debug "Backup Running"
Wend
WriteLog("Backup","Database Backup Successfully saved to "+MyLocation)
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
  If FindString(output$,"Backup OE Classic",1,#PB_String_NoCase)
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
Protected GetDate.s, myid
 If GetPidProcessEx("OEClassic.exe")
   MessageRequester("Error","OE Classic must be closed before backing up.",#MB_ICONERROR)
    WriteLog("Backup","Backup failed, OE Classic was not closed.")
 Else
GetDate=FormatDate("%yyyy%mm%dd",Date())
  While WaitWindowEvent(1):Wend
StatusBarText(#StatusBar,0,"Please wait.....creating backup",#PB_StatusBar_Center)
 If FindString(MyLocation," ",1)
   myid=RunProgram("7z.exe","a -mx=9 "+Chr(34)+MyLocation+Chr(34)+"OEClassicBackup_"+GetDate+".7z "+oepath,"",#PB_Program_Open|#PB_Program_Hide)
 Else
   myid=RunProgram("7z.exe","a -mx=9 "+MyLocation+"OEClassicBackup_"+GetDate+".7z "+oepath,"",#PB_Program_Open|#PB_Program_Hide)
 EndIf
While ProgramRunning(myid)
Debug "Backup Running"
Wend
StatusBarText(#StatusBar,0,"Ready",#PB_StatusBar_Center)
WriteLog("Backup","Database Backup Successfully saved to "+MyLocation)
MessageRequester("Success","Backup Completed",#MB_ICONINFORMATION)
EndIf
EndProcedure

Procedure CreateBackupTask()
OpenFile(0,"MyTask.xml",#PB_Ascii)
WriteStringN(0,"<?xml version="+Chr(34)+"1.0"+Chr(34)+" encoding="+Chr(34)+"UTF-16"+Chr(34)+"?>",#PB_Ascii)
WriteStringN(0,"<Task version="+Chr(34)+"1.2"+Chr(34)+" xmlns="+Chr(34)+"http://schemas.microsoft.com/windows/2004/02/mit/task"+Chr(34)+">",#PB_Ascii)
WriteStringN(0,"<RegistrationInfo>",#PB_Ascii)
WriteStringN(0,"<Date>2021-02-21T10:54:55.4869081</Date>",#PB_Ascii)
WriteStringN(0,"<Author>"+GetEnvironmentVariable("USERNAME")+"</Author>",#PB_Ascii)
WriteStringN(0,"<URI>\Backup OE Classic</URI>",#PB_Ascii)
WriteStringN(0,"</RegistrationInfo>",#PB_Ascii)
WriteStringN(0,"<Triggers>",#PB_Ascii)
WriteStringN(0,"<CalendarTrigger>",#PB_Ascii)
WriteStringN(0,"<StartBoundary>2021-02-26T23:59:00-06:00</StartBoundary>",#PB_Ascii)
WriteStringN(0,"<ExecutionTimeLimit>PT8H</ExecutionTimeLimit>",#PB_Ascii)
WriteStringN(0,"<Enabled>true</Enabled>",#PB_Ascii)
WriteStringN(0,"<ScheduleByWeek>",#PB_Ascii)
WriteStringN(0,"<DaysOfWeek>",#PB_Ascii)
WriteStringN(0,"<Friday />",#PB_Ascii)
WriteStringN(0,"</DaysOfWeek>",#PB_Ascii)
WriteStringN(0,"<WeeksInterval>1</WeeksInterval>",#PB_Ascii)
WriteStringN(0,"</ScheduleByWeek>",#PB_Ascii)
WriteStringN(0,"</CalendarTrigger>",#PB_Ascii)
WriteStringN(0,"</Triggers>",#PB_Ascii)
WriteStringN(0,"<Principals>",#PB_Ascii)
WriteStringN(0,"<Principal id="+Chr(34)+"Author"+Chr(34)+">",#PB_Ascii)
WriteStringN(0,"<UserId>"+GetEnvironmentVariable("USERNAME")+"</UserId>",#PB_Ascii)
WriteStringN(0,"<LogonType>InteractiveToken</LogonType>")
WriteStringN(0,"<RunLevel>HighestAvailable</RunLevel>",#PB_Ascii)
WriteStringN(0,"</Principal>",#PB_Ascii)
WriteStringN(0,"</Principals>",#PB_Ascii)
WriteStringN(0,"<Settings>",#PB_Ascii)
WriteStringN(0,"<MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>",#PB_Ascii)
WriteStringN(0,"<DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>",#PB_Ascii)
WriteStringN(0,"<StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>",#PB_Ascii)
WriteStringN(0,"<AllowHardTerminate>true</AllowHardTerminate>",#PB_Ascii)
WriteStringN(0,"<StartWhenAvailable>false</StartWhenAvailable>",#PB_Ascii)
WriteStringN(0,"<RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>",#PB_Ascii)
WriteStringN(0,"<IdleSettings>",#PB_Ascii)
WriteStringN(0,"<StopOnIdleEnd>true</StopOnIdleEnd>",#PB_Ascii)
WriteStringN(0,"<RestartOnIdle>false</RestartOnIdle>",#PB_Ascii)
WriteStringN(0,"</IdleSettings>",#PB_Ascii)
WriteStringN(0,"<AllowStartOnDemand>true</AllowStartOnDemand>",#PB_Ascii)
WriteStringN(0,"<Enabled>true</Enabled>",#PB_Ascii)
WriteStringN(0,"<Hidden>false</Hidden>",#PB_Ascii)
WriteStringN(0,"<RunOnlyIfIdle>false</RunOnlyIfIdle>",#PB_Ascii)
WriteStringN(0,"<WakeToRun>false</WakeToRun>",#PB_Ascii)
WriteStringN(0,"<ExecutionTimeLimit>PT8H</ExecutionTimeLimit>",#PB_Ascii)
WriteStringN(0,"<Priority>7</Priority>",#PB_Ascii)
WriteStringN(0,"</Settings>",#PB_Ascii)
WriteStringN(0,"<Actions Context="+Chr(34)+"Author"+Chr(34)+">",#PB_Ascii)
WriteStringN(0,"<Exec>",#PB_Ascii)
WriteStringN(0,"<Command>"+GetCurrentDirectory()+"OEClassicBackup.exe</Command>",#PB_Ascii)
WriteStringN(0,"<Arguments>/b</Arguments>",#PB_Ascii)
WriteStringN(0,"<WorkingDirectory>"+GetCurrentDirectory()+"</WorkingDirectory>",#PB_Ascii)
WriteStringN(0,"</Exec>",#PB_Ascii)
WriteStringN(0,"</Actions>",#PB_Ascii)
WriteStringN(0,"</Task>",#PB_Ascii)
CloseFile(0)
Delay(500)
RunProgram("schtasks","/create /tn "+Chr(34)+"Backup OE Classic"+Chr(34)+" /xml mytask.xml","",#PB_Program_Hide|#PB_Program_Wait)
  WriteLog("Backup","Scheduled task successfully created.")
   DeleteFile("MyTask.xml",#PB_FileSystem_Force)
    OpenPreferences("oebackup.prefs")
     WritePreferenceInteger("SchTsk",1)
    ClosePreferences()
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
   MessageRequester("Error","OE Classic must be closed before restoring data.",#MB_ICONERROR)
    WriteLog("Backup","Restore failed, OEClassic was not closed.")
 Else
   myrestore=MessageRequester("Restore","Are you sure you wish to restore from a backup?",#PB_MessageRequester_YesNo|#MB_ICONQUESTION)
  If myrestore=#PB_MessageRequester_Yes
   If MyLocation<>""
     startpath.s=MyLocation
   Else
     startpath.s=""
   EndIf
     BackupFile=OpenFileRequester("Restore Backup",startpath,"7-Zip (*.7z)|*.7z",0)
    If BackupFile<>""
      yesrestore=MessageRequester("Warning","This will overwrite your current OE Classic data."+#CRLF$+"Are you sure you wish to continue?",#PB_MessageRequester_YesNo|#MB_ICONEXCLAMATION)
     If yesrestore=#PB_MessageRequester_Yes
       restorepath.s=GetEnvironmentVariable("USERPROFILE")
        restorepath+"\AppData\Local\"
         DeleteDirectory(oepath,"*.*",#PB_FileSystem_Recursive|#PB_FileSystem_Force)
          StatusBarText(#StatusBar,0,"Please Wait.....restoring data",#PB_StatusBar_Center)
      If FindString(BackupFile," ",1)
        myid=RunProgram("7z.exe","x -o"+restorepath+" "+Chr(34)+BackupFile+Chr(34),"",#PB_Program_Open|#PB_Program_Hide)
      Else
        myid=RunProgram("7z.exe","x -o"+restorepath+" "+BackupFile,"",#PB_Program_Open|#PB_Program_Hide)
      EndIf
       While WaitWindowEvent(1):Wend
        While ProgramRunning(myid)
          Debug "Restore Running"
        Wend
        StatusBarText(#Statusbar,0,"Ready",#PB_StatusBar_Center)
       WriteLog("Backup","Restore Successful")
      MessageRequester("Success","Restore Completed",#MB_ICONINFORMATION)
     Else
       MessageRequester("Cancelled","Your restore request has been cancelled.",#MB_ICONINFORMATION)
     EndIf
    EndIf
  Else
    MessageRequester("Cancelled","Your restore request has been cancelled.",#MB_ICONINFORMATION)
  EndIf
 EndIf
EndProcedure

Procedure UpdatePrefs()
If GetGadgetState(#Checkbox_CloseToTray)=#PB_Checkbox_Checked
  OpenPreferences("oebackup.prefs")
   WritePreferenceInteger("CloseToTray",1)
  ClosePreferences()
Else
  OpenPreferences("oebackup.prefs")
   WritePreferenceInteger("CloseToTray",0)
  ClosePreferences()
EndIf

 If GetGadgetState(#Checkbox_RememberBackupLocation)=#PB_Checkbox_Checked
   OpenPreferences("oebackup.prefs")
    WritePreferenceString("BkUpDir",GetGadgetText(#String_BackupPath))
    WritePreferenceInteger("RememberDir",1)
   ClosePreferences()
 Else
   OpenPreferences("oebackup.prefs")
    WritePreferenceString("BkUpDir","")
    WritePreferenceInteger("RememberDir",0)
   ClosePreferences()
 EndIf

  If GetGadgetState(#Checkbox_StartInTray)=#PB_Checkbox_Checked
    OpenPreferences("oebackup.prefs")
     WritePreferenceInteger("OpenToTray",1)
    ClosePreferences()
  Else
    OpenPreferences("oebackup.prefs")
     WritePreferenceInteger("OpenToTray",0)
    ClosePreferences()
  EndIf

   If GetGadgetState(#Checkbox_StartWithLogon)=#PB_Checkbox_Checked
     OpenPreferences("oebackup.prefs")
      WritePreferenceInteger("AutoStart",1)
     ClosePreferences()
      CreateShortcut(GetCurrentDirectory()+"oeclassicbackup.exe",GetEnvironmentVariable("USERPROFILE")+"\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\OEClassicBackup.lnk","","OE Classic Backup",GetCurrentDirectory(),#SW_SHOWNORMAL,GetCurrentDirectory()+"oeclassicbackup.exe",0)
   Else
     DeleteFile(GetEnvironmentVariable("USERPROFILE")+"\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\OEClassicBackup.lnk",#PB_FileSystem_Force)
   EndIf
EndProcedure
;}

;{ Create Preference File
If OpenPreferences("oebackup.prefs")=0
  CreatePreferences("oebackup.prefs")
   OpenPreferences("oebackup.prefs")
    WritePreferenceString("BkUpDir","")
    WritePreferenceInteger("RememberDir",0)
    WritePreferenceInteger("CloseToTray",0)
    WritePreferenceInteger("OpenToTray",0)
    WritePreferenceInteger("AutoStart",0)
    WritePreferenceInteger("SchTsk",0)
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
KeepDir=ReadPreferenceInteger("RememberDir",0)
CloseToTray=ReadPreferenceInteger("CloseToTray",0)
OpenToTray=ReadPreferenceInteger("OpenToTray",0)
AutoStart=ReadPreferenceInteger("AutoStart",0)
SchedTask=ReadPreferenceInteger("SchTsk",0)
ClosePreferences()
;}

;{ Create Window and Gadgets
If FileSize(GetEnvironmentVariable("userprofile")+"\Appdata\Local\OEClassic")<>-1
  oepath.s=GetEnvironmentVariable("userprofile")+"\Appdata\Local\OEClassic"
Else
  MessageRequester("Error","OE Classic path not found!!!",#MB_ICONERROR)
   End
EndIf
If ProgramParameter()=""
 If OpenToTray=1
   startme=#PB_Window_Invisible|#PB_Window_ScreenCentered
 Else
   startme=#PB_Window_ScreenCentered
 EndIf
  OpenWindow(#Window_Main,0,0,500,200,"OE Classic Backup",startme|#PB_Window_SystemMenu)
   PanelGadget(#Panel1,5,5,492,170)
    AddGadgetItem(#Panel1,0,"Main")
     TextGadget(#Text_Path,225,7,100,20,"OE Classic Path:")
      StringGadget(#String_Path,60,25,400,20,oepath)
       HyperLinkGadget(#Text_BackupPath,225,47,100,20,"BackUp Location:",#Blue)
       GadgetToolTip(#Text_BackupPath,"Click to clear backup path")
        StringGadget(#String_BackupPath,60,67,400,20,MyLocation)
         ButtonGadget(#Button_SetBackupPath,175,105,150,30,"Set Backup Location")
         ButtonGadget(#Button_Backup,20,105,150,30,"Create Backup")
         ButtonGadget(#Button_OpenBackupLocation,175,105,150,30,"Open Backup Location")
         ButtonGadget(#Button_RestoreBackup,330,105,150,30,"Restore Backup")
          CreateStatusBar(#StatusBar,WindowID(#Window_Main))
           AddStatusBarField(500)
           StatusBarText(#StatusBar,0,"Ready",#PB_StatusBar_Center)
    CloseGadgetList()
    OpenGadgetList(#Panel1)
     AddGadgetItem(#Panel1,1,"Options")
      CheckBoxGadget(#Checkbox_RememberBackupLocation,20,20,165,20,"Remember Backup Location")
      CheckBoxGadget(#Checkbox_CloseToTray,20,45,105,20,"Close to Tray")
      CheckBoxGadget(#Checkbox_StartInTray,20,70,185,20,"Start Application in System Tray")
      CheckBoxGadget(#Checkbox_StartWithLogon,20,95,110,20,"Start With Logon")
       ButtonGadget(#Button_CreateDesktopIcon,19,115,165,20,"Create Desktop Shortcut")
       ButtonGadget(#Button_SetTask,190,115,165,20,"Create Backup Task")
    CloseGadgetList()
     StickyWindow(#Window_Main,#True)
  If CreatePopupImageMenu(#Menu_SysTray);System Tray Menu
    MenuItem(#RestoreApp,"Restore Application Window",CatchImage(#Icon_RestoreApp,?Icon_RestoreApp))
    MenuBar()
     MenuItem(#BackupDB,"Backup Database",CatchImage(#Icon_BackupDB,?Icon_BackupDB))
      MenuItem(#RestoreDB,"Restore Database",CatchImage(#Icon_RestoreDB,?Icon_RestoreDB))
       MenuItem(#OpenBackupFolder,"Open Backup Folder",CatchImage(#Icon_BackupFolder,?Icon_BackupFolder))
    MenuBar()
        MenuItem(#QuitApp,"Quit",CatchImage(#Icon_Quit,?Icon_Quit))
  EndIf
   If KeepDir=1
     SetGadgetState(#Checkbox_RememberBackupLocation,#PB_Checkbox_Checked)
   EndIf
    If CloseToTray=1
      SetGadgetState(#Checkbox_CloseToTray,#PB_Checkbox_Checked)
    EndIf
     If OpenToTray=1
       SetGadgetState(#Checkbox_StartInTray,#PB_Checkbox_Checked)
        systrayicon.l = CatchImage(#Icon_SysTray, ?Icon_Systray)
         AddSysTrayIcon(0,WindowID(#Window_Main),systrayicon)
          SysTrayIconToolTip(0,"OE Classic Backup - Click for Options")
           ShowWindow_(WindowID(#Window_Main),#SW_HIDE)
     EndIf
      If AutoStart=1
       SetGadgetState(#Checkbox_StartWithLogon,#PB_Checkbox_Checked)
      EndIf
       If SchedTask=1
         HideGadget(#Button_SetTask,1)
       EndIf
EndIf
;}

;{ Main Loop
Repeat

;{ Disable/Hide Gadgets
If FileSize(GetEnvironmentVariable("userprofile")+"\Desktop\OE Classic Backup.lnk")<>-1
  HideGadget(#Button_CreateDesktopIcon,1)
EndIf

If MyLocation<>""
  HideGadget(#Button_SetBackupPath,1)
   HideGadget(#Button_OpenBackupLocation,0)
    DisableGadget(#Button_Backup,0)
     DisableGadget(#Button_OpenBackupLocation,0)
      DisableMenuItem(#Menu_SysTray,#BackupDB,0)
       DisableMenuItem(#Menu_SysTray,#RestoreDB,0)
        DisableMenuItem(#Menu_SysTray,#OpenBackupFolder,0)
Else
  HideGadget(#Button_SetBackupPath,0)
   HideGadget(#Button_OpenBackupLocation,1)
    DisableGadget(#Button_Backup,1)
     DisableGadget(#Button_OpenBackupLocation,1)
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
  
      Case #Button_Backup
        CreateBackup()

      Case #Button_CreateDesktopIcon
        CreateShortcut(GetCurrentDirectory()+"oeclassicbackup.exe",GetEnvironmentVariable("USERPROFILE")+"\Desktop\OE Classic Backup.lnk","","OE Classic Backup",GetCurrentDirectory(),#SW_SHOWNORMAL,GetCurrentDirectory()+"oeclassicbackup.exe",0)

      Case #Button_OpenBackupLocation
        OpenBackupLocation()

      Case #Button_RestoreBackup
        RestoreBackup()

      Case #Button_SetBackupPath
        MyLocation=PathRequester("Choose Backup Location","C:\OEBackup")
         If MyLocation<>""
           SetGadgetText(#String_BackupPath,MyLocation)
         EndIf

       Case #Button_SetTask
         CreateBackupTask()
          If CheckforTask()=1
            HideGadget(#Button_SetTask,1)
          EndIf

      Case #Checkbox_CloseToTray
        UpdatePrefs()

      Case #Checkbox_RememberBackupLocation
        UpdatePrefs()

      Case #Checkbox_StartInTray
        UpdatePrefs()

      Case #Checkbox_StartWithLogon
        UpdatePrefs()

      Case #Text_BackupPath
        SetGadgetText(#String_BackupPath,"")
         SetGadgetState(#Checkbox_RememberBackupLocation,#PB_Checkbox_Unchecked)
          MyLocation=""

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
           UpdatePrefs()
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
            SysTrayIconToolTip(0,"OE Classic Backup - Click for Options")
             ShowWindow_(WindowID(#Window_Main),#SW_HIDE)
        Else
          End
        EndIf
;}
EndSelect

ForEver
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
; IDE Options = PureBasic 5.73 LTS (Windows - x64)
; CursorPosition = 661
; FirstLine = 50
; Folding = AJhAg
; EnableThread
; EnableXP
; EnableAdmin
; UseIcon = gfx\icon3.ico
; Executable = C:\Temp\OEClassicBackup.exe
; Compiler = PureBasic 5.73 LTS (Windows - x86)