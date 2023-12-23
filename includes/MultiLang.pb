; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; Project name : Basic Language Management System
; File Name : Basic Language Management System.pb
; File version: 1.0.0
; Programming : OK
; Programmed by : Guimauve
; Date : 09-12-2012
; Last Update : 09-12-2012
; PureBasic code : V5.00
; Platform : Windows, Linux, MacOS X
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; Programming notes
;
; This is a little remake of Freak's code posted in English 
; forum. It also include an extra SaveLangauge procedure 
; added by Thyphoon.
;
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< Déclaration de la Structure <<<<<

Structure LanguageGroup
  
  Name.s
  GroupStart.l
  GroupEnd.l
  IndexTable.l[256]
  
EndStructure

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< Les observateurs <<<<<

Macro GetLanguageGroupName(LanguageGroupA)
  
  LanguageGroupA\Name
  
EndMacro

Macro GetLanguageGroupGroupStart(LanguageGroupA)
  
  LanguageGroupA\GroupStart
  
EndMacro

Macro GetLanguageGroupGroupEnd(LanguageGroupA)
  
  LanguageGroupA\GroupEnd
  
EndMacro

Macro GetLanguageGroupIndexTable(LanguageGroupA, IndexTableID)
  
  LanguageGroupA\IndexTable[IndexTableID]
  
EndMacro

; <<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< Les mutateurs <<<<<

Macro SetLanguageGroupName(LanguageGroupA, P_Name)
  
  GetLanguageGroupName(LanguageGroupA) = P_Name
  
EndMacro

Macro SetLanguageGroupGroupStart(LanguageGroupA, P_GroupStart)
  
  GetLanguageGroupGroupStart(LanguageGroupA) = P_GroupStart
  
EndMacro

Macro SetLanguageGroupGroupEnd(LanguageGroupA, P_GroupEnd)
  
  GetLanguageGroupGroupEnd(LanguageGroupA) = P_GroupEnd
  
EndMacro

Macro SetLanguageGroupIndexTable(LanguageGroupA, IndexTableID, P_IndexTable)
  
  GetLanguageGroupIndexTable(LanguageGroupA, IndexTableID) = P_IndexTable
  
EndMacro

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< L'opérateur Reset <<<<<

Macro ResetLanguageGroup(LanguageGroupA)
  
  SetLanguageGroupName(LanguageGroupA, "")
  SetLanguageGroupGroupStart(LanguageGroupA, 0)
  SetLanguageGroupGroupEnd(LanguageGroupA, 0)
  
  For IndexTableID = 0 To 255
    SetLanguageGroupIndexTable(LanguageGroupA, IndexTableID, 0)
  Next
  
EndMacro

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< Code généré en : 00.004 secondes (25250.00 lignes/seconde) <<<<<
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< Déclaration de la Structure <<<<<

Structure Language
  
  GroupCount.l ; Special : Increment
  StringCount.l ; Special : Increment
  Array Groups.LanguageGroup(0)
  Array Strings.s(0)
  Array Names.s(0)
  
EndStructure

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< Les observateurs <<<<<

Macro GetLanguageGroupCount(LanguageA)
  
  LanguageA\GroupCount
  
EndMacro

Macro GetLanguageStringCount(LanguageA)
  
  LanguageA\StringCount
  
EndMacro

Macro GetLanguageGroups(LanguageA)
  
  LanguageA\Groups()
  
EndMacro

Macro GetLanguageStrings(LanguageA)
  
  LanguageA\Strings()
  
EndMacro

Macro GetLanguageNames(LanguageA)
  
  LanguageA\Names()
  
EndMacro

; <<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< Les mutateurs <<<<<

Macro SetLanguageGroupCount(LanguageA, P_GroupCount)
  
  GetLanguageGroupCount(LanguageA) = P_GroupCount
  
EndMacro

Macro SetLanguageStringCount(LanguageA, P_StringCount)
  
  GetLanguageStringCount(LanguageA) = P_StringCount
  
EndMacro

Macro SetLanguageGroups(LanguageA, P_Groups)
  
  CopyLanguageGroups(P_Groups, GetLanguageGroups(LanguageA))
  
EndMacro

Macro SetLanguageStrings(LanguageA, P_Strings)
  
  GetLanguageStrings(LanguageA) = P_Strings
  
EndMacro

Macro SetLanguageNames(LanguageA, P_Names)
  
  GetLanguageNames(LanguageA) = P_Names
  
EndMacro

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< Les macros complémentaires pour les Tableaux dynamiques <<<<<

Macro GetLanguageGroupsElement(LanguageA, Groups_ID_1)
  
  LanguageA\Groups(Groups_ID_1)
  
EndMacro

Macro SetLanguageGroupsElement(LanguageA, Groups_ID_1, P_Element)
  
  GetLanguageGroupsElement(LanguageA, Groups_ID_1) = P_Element
  
EndMacro

Macro ReDimLanguageGroups(LanguageA, Groups_Max_1D)
  
  ReDim GetLanguageGroupsElement(LanguageA, Groups_Max_1D)
  
EndMacro

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< Les macros complémentaires pour les Tableaux dynamiques <<<<<

Macro GetLanguageStringsElement(LanguageA, Strings_ID_1)
  
  LanguageA\Strings(Strings_ID_1)
  
EndMacro

Macro SetLanguageStringsElement(LanguageA, Strings_ID_1, P_Element)
  
  GetLanguageStringsElement(LanguageA, Strings_ID_1) = P_Element
  
EndMacro

Macro ReDimLanguageStrings(LanguageA, Strings_Max_1D)
  
  ReDim GetLanguageStringsElement(LanguageA, Strings_Max_1D)
  
EndMacro

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< Les macros complémentaires pour les Tableaux dynamiques <<<<<

Macro GetLanguageNamesElement(LanguageA, Names_ID_1)
  
  LanguageA\Names(Names_ID_1)
  
EndMacro

Macro SetLanguageNamesElement(LanguageA, Names_ID_1, P_Element)
  
  GetLanguageNamesElement(LanguageA, Names_ID_1) = P_Element
  
EndMacro

Macro ReDimLanguageNames(LanguageA, Names_Max_1D)
  
  ReDim GetLanguageNamesElement(LanguageA, Names_Max_1D)
  
EndMacro

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< Les opérateurs spéciaux <<<<<

Macro IncrementLanguageGroupCount(LanguageA, P_Increment = 1)
  
  SetLanguageGroupCount(LanguageA, GetLanguageGroupCount(LanguageA) + P_Increment)
  
EndMacro

Macro IncrementLanguageStringCount(LanguageA, P_Increment = 1)
  
  SetLanguageStringCount(LanguageA, GetLanguageStringCount(LanguageA) + P_Increment)
  
EndMacro

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< L'opérateur Reset <<<<<

Macro ResetLanguage(LanguageA)
  
  SetLanguageGroupCount(LanguageA, 0)
  SetLanguageStringCount(LanguageA, 0)
  
  For Groups_ID_1 = 0 To ArraySize(GetLanguageGroups(LanguageA), 1)
    ResetLanguageGroup(GetLanguageGroupsElement(LanguageA, Groups_ID_1))
  Next
  
  For Strings_ID_1 = 0 To ArraySize(GetLanguageStrings(LanguageA), 1)
    SetLanguageStringsElement(LanguageA, Strings_ID_1, "")
    SetLanguageNamesElement(LanguageA, Strings_ID_1, "")
  Next
  
  ClearStructure(LanguageA, Language)
  InitializeStructure(LanguageA, Language)
  
EndMacro

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< Code généré en : 00.009 secondes (25000.00 lignes/seconde) <<<<<
; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< The LoadLanguage operator <<<<<

Procedure LoadLanguage(*LanguageA.Language, P_FileName.s = "")
  
  Protected GroupID, StringID.l, CharID.a, FirstChar.a
  
  ResetLanguage(*LanguageA)
  
  ; do a quick count in the datasection first:
  
	Restore Language
	Repeat

		Read.s GroupName.s
		Read.s String.s

		GroupName = UCase(GroupName)

		If GroupName = "_GROUP_"
			IncrementLanguageGroupCount(*LanguageA)
		ElseIf GroupName = "_END_"
			Break
		Else
			IncrementLanguageStringCount(*LanguageA)
		EndIf
		
	ForEver
	
	ReDimLanguageGroups(*LanguageA, GetLanguageGroupCount(*LanguageA))
	ReDimLanguageStrings(*LanguageA, GetLanguageStringCount(*LanguageA))
	ReDimLanguageNames(*LanguageA, GetLanguageStringCount(*LanguageA))

	; Now load the standard language:
 
	GroupID = 0
	StringID = 0  

	Restore Language
	Repeat

		Read.s GroupName.s
		Read.s String.s

		GroupName = UCase(GroupName)

		If GroupName = "_GROUP_"
		  SetLanguageGroupGroupEnd(GetLanguageGroupsElement(*LanguageA, GroupID), StringID)
			GroupID + 1
			SetLanguageGroupName(GetLanguageGroupsElement(*LanguageA, GroupID), UCase(String))
			SetLanguageGroupGroupStart(GetLanguageGroupsElement(*LanguageA, GroupID), StringID + 1)

			For IndexTableID = 0 To 255
			  SetLanguageGroupIndexTable(GetLanguageGroupsElement(*LanguageA, GroupID), IndexTableID, 0)
			Next	

		ElseIf GroupName = "_END_"
			Break

		Else
		  StringID + 1
		  SetLanguageNamesElement(*LanguageA, StringID, GroupName + Chr(1) + String ); keep name and string together for easier sorting

		EndIf
		
	ForEver

	SetLanguageGroupGroupEnd(GetLanguageGroupsElement(*LanguageA, GroupID), StringID) ; set end for the last group!
	
	; Now do the sorting and the indexing for each group
	
	For GroupID = 1 To GetLanguageGroupCount(*LanguageA)
	  
		If GetLanguageGroupGroupStart(GetLanguageGroupsElement(*LanguageA, GroupID)) <= GetLanguageGroupGroupEnd(GetLanguageGroupsElement(*LanguageA, GroupID))  ; sanity check.. check for empty groups
			
			SortArray(GetLanguageNames(*LanguageA), 0, GetLanguageGroupGroupStart(GetLanguageGroupsElement(*LanguageA, GroupID)), GetLanguageGroupGroupEnd(GetLanguageGroupsElement(*LanguageA, GroupID)))
	
			CharID = 0
			
			For StringID = GetLanguageGroupGroupStart(GetLanguageGroupsElement(*LanguageA, GroupID)) To GetLanguageGroupGroupEnd(GetLanguageGroupsElement(*LanguageA, GroupID))

			  SetLanguageStringsElement(*LanguageA, StringID, StringField(GetLanguageNamesElement(*LanguageA, StringID), 2, Chr(1)))
			  SetLanguageNamesElement(*LanguageA, StringID, StringField(GetLanguageNamesElement(*LanguageA, StringID), 1, Chr(1)))
			  FirstChar = Asc(Left(GetLanguageNamesElement(*LanguageA, StringID), 1))
			  
				If FirstChar <> CharID
				  CharID = FirstChar
				  SetLanguageGroupIndexTable(GetLanguageGroupsElement(*LanguageA, GroupID), CharID, StringID)
				EndIf
				
			Next
			
		EndIf
		
	Next

	; Now try to load an external language file
    
	If P_FileName <> ""
			
	  If OpenPreferences(P_FileName)
	    
	    For GroupID = 1 To GetLanguageGroupCount(*LanguageA)
	      
				If GetLanguageGroupGroupStart(GetLanguageGroupsElement(*LanguageA, GroupID)) <= GetLanguageGroupGroupEnd(GetLanguageGroupsElement(*LanguageA, GroupID))  ; sanity check.. check for empty groups
				  
				  PreferenceGroup(GetLanguageGroupName(GetLanguageGroupsElement(*LanguageA, GroupID)))
					
					For StringID = GetLanguageGroupGroupStart(GetLanguageGroupsElement(*LanguageA, GroupID)) To GetLanguageGroupGroupEnd(GetLanguageGroupsElement(*LanguageA, GroupID))
					  SetLanguageStringsElement(*LanguageA, StringID, ReadPreferenceString(GetLanguageNamesElement(*LanguageA, StringID), GetLanguageStringsElement(*LanguageA, StringID)))
					Next
					
				EndIf
				
			Next 
			
			ClosePreferences()   
			
			ProcedureReturn #True
		EndIf    

	EndIf
	
	ProcedureReturn #True
EndProcedure

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< The SaveLanguage operator <<<<<

Procedure SaveLanguage(*LanguageA.Language, P_FileName.s)
  
  Protected GroupID, StringID.l
  
  If CreatePreferences(P_FileName)
    
    For GroupID = 1 To GetLanguageGroupCount(*LanguageA)
      
      PreferenceGroup(GetLanguageGroupName(GetLanguageGroupsElement(*LanguageA, GroupID)))
      
      For StringID = GetLanguageGroupGroupStart(GetLanguageGroupsElement(*LanguageA, GroupID)) To GetLanguageGroupGroupEnd(GetLanguageGroupsElement(*LanguageA, GroupID))
        WritePreferenceString(GetLanguageNamesElement(*LanguageA, StringID), GetLanguageStringsElement(*LanguageA, StringID))
      Next
      
    Next
    
    ClosePreferences()
    
  EndIf
  
EndProcedure

; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; <<<<< Get Language text <<<<<

Procedure.s Language(*LanguageA.Language, P_Group.s, P_Name.s)
  
  Static GroupID.l  ; for quicker access when using the same group more than once
  Protected String.s, StringID, Result
  
  P_Group  = UCase(P_Group)
  P_Name   = UCase(P_Name)
  String = "##### String not found! #####"  ; to help find bugs
  
  If UCase(GetLanguageGroupName(GetLanguageGroupsElement(*LanguageA, GroupID))) <> P_Group  ; check if it is the same group as last time
    
    For GroupID = 1 To GetLanguageGroupCount(*LanguageA)
      If P_Group = GetLanguageGroupName(GetLanguageGroupsElement(*LanguageA, GroupID))
        Break
      EndIf
    Next 
    
    If GroupID > GetLanguageGroupCount(*LanguageA)  ; check if group was found
      GroupID = 0
    EndIf
    
  EndIf
  
  If GroupID <> 0
    
    StringID = GetLanguageGroupIndexTable(GetLanguageGroupsElement(*LanguageA, GroupID), Asc(Left(P_Name, 1)))

    If StringID <> 0
      
      Repeat
        
        Result = CompareMemoryString(@P_Name, @GetLanguageNamesElement(*LanguageA, StringID))
        
        If Result = 0
          String = GetLanguageStringsElement(*LanguageA, StringID)
          Break
        ElseIf Result = -1 ; string not found!
          Break
        EndIf
        
        StringID + 1
        
      Until StringID > GetLanguageGroupGroupEnd(GetLanguageGroupsElement(*LanguageA, GroupID))
      
    EndIf
    
  EndIf
  
  ProcedureReturn String
EndProcedure

; ; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
; ; <<<<< DataSection pour les textes par défaut <<<<<
; 
; DataSection
;   
;   ; Here the default language is specified. It is a list of Group, Name pairs,
;   ; with some special keywords for the Group:
;   ;
;   ; "_GROUP_" will indicate a new group in the datasection, the second value is the group name
;   ; "_END_" will indicate the end of the language list (as there is no fixed number)
;   ;
;   ; Note: The identifier strings are case insensitive to make live easier :)
;   
;   Language:
;   Data.s "_GROUP_",            "WindowTitle"
; 
;   Data.s "WinTitle",           "OE Classic Backup"
;   ; ===================================================
;   Data.s "_GROUP_",            "MenuTitle"
;   ; ===================================================
;   
;   Data.s "File",             "File"
;   Data.s "Edit",             "Edit"
;   
;   ; ===================================================
;   Data.s "_GROUP_",            "MenuItem"
;   ; ===================================================
;   
;   Data.s "New",              "New"
;   Data.s "Open",             "Open..."
;   Data.s "Save",             "Save"
;   
;   ; ===================================================
;   Data.s "_END_",              ""
;   ; ===================================================
;   
; EndDataSection
; 
; ; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< 
; ; <<<<< !!! WARNING - YOU ARE NOW IN A TESTING ZONE - WARNING !!! <<<<< 
; ; <<<<< !!! WARNING - THIS CODE SHOULD BE COMMENTED - WARNING !!! <<<<< 
; ; <<<<< !!! WARNING - BEFORE THE FINAL COMPILATION. - WARNING !!! <<<<< 
; ; <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< 
; 
; Global Language.Language
; 
; LoadLanguage(Language)                ; load default language
; LoadLanguage(Language.Language, "lang\german.prefs") ; uncomment this to load the german file
; 
; ; get some language strings
; Debug Language(Language, "WindowTitle", "WinTitle")
; Debug Language(Language, "MenuTitle", "Edit")
; Debug Language(Language, "MenuItem", "Open")
; 
; OpenWindow(0,0,0,250,250,Language(Language,"WindowTitle","WinTitle"),#PB_Window_ScreenCentered|#PB_Window_SystemMenu)
; Repeat
; event=WaitWindowEvent(1)
; Select event
;   Case #PB_Event_CloseWindow
;     End
; EndSelect
; ForEver
; 
; ResetLanguage(Language)

; <<<<<<<<<<<<<<<<<<<<<<<
; <<<<< END OF FILE <<<<<
; <<<<<<<<<<<<<<<<<<<<<<<
; IDE Options = PureBasic 6.04 LTS (Windows - x64)
; CursorPosition = 557
; FirstLine = 179
; Folding = AAAAAw
; EnableXP
; DPIAware