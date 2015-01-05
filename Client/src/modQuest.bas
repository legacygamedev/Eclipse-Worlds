Attribute VB_Name = "modQuest"
Option Explicit


Public Const QUESTICON_LENGTH = 5

Public Const QUESTNAME_LENGTH = 40

Public Const QUESTDESC_LENGTH = 300

Public QuestRequest                    As Long

Public QuestLogDisplaySlot             As Long

Public QuestLogQuest                   As Long

Public QuestLog()       As Long

Public Quest()          As QuestRec

Public PlayerQuest()   As PlayerQuestRec

'Constants to use for tasks
Public Const TASK_KILL                 As Byte = 1

Public Const TASK_GATHER               As Byte = 2

Public Const TASK_VARIABLE             As Byte = 3

Public Const TASK_GETSKILL             As Byte = 4

Public Const ACTION_GIVE_ITEM          As Byte = 5

Public Const ACTION_TAKE_ITEM          As Byte = 6

Public Const ACTION_SHOWMSG            As Byte = 7

Public Const ACTION_ADJUST_LVL         As Byte = 8

Public Const ACTION_ADJUST_EXP         As Byte = 9

Public Const ACTION_WARP               As Byte = 10

Public Const ACTION_ADJUST_STAT_LVL    As Byte = 11

Public Const ACTION_ADJUST_SKILL_LVL   As Byte = 12

Public Const ACTION_ADJUST_SKILL_EXP   As Byte = 13

Public Const ACTION_ADJUST_STAT_POINTS As Byte = 14

Public Const ACTION_SETVARIABLE        As Byte = 15

Public Const ACTION_PLAYSOUND          As Byte = 16

Public Const Quest_Icon_Start          As Byte = 1

Public Const Quest_Icon_Finished   As Byte = 2

Public Const Quest_Icon_Progress       As Byte = 3

Public Const Quest_Icon_Completed      As Byte = 4

'Constants for list mover
Public Const LIST_CLI                  As Byte = 1

Public Const LIST_TASK                 As Byte = 2

Private Type QuestAmountRec

    ID() As Integer

End Type

Private Type PlayerQuestRec

    QuestCompleted() As Byte
    QuestCLI() As Integer
    QuestTask() As Integer
    QuestAmount() As QuestAmountRec

End Type

Private Type RequirementsRec

    AccessReq As Long
    LevelReq As Long
    GenderReq As Long
    ClassReq As Long
    SkillReq As Long
    SkillLevelReq As Long
    Stat_Req(1 To Stats.Stat_Count - 1) As Long

End Type

Private Type ActionRec

    TextHolder As String * QUESTDESC_LENGTH
    ActionID As Byte
    amount As Long
    MainData As Long
    SecondaryData As Long
    TertiaryData As Long
    QuadData As Long

End Type

Private Type CLIRec

    ItemIndex As Long
    isNPC As Long
    Max_Actions As Long
    Action() As ActionRec

End Type

Private Type QuestRec

    Name As String * QUESTNAME_LENGTH
    Description As String * QUESTDESC_LENGTH
    CanBeRetaken As Byte
    Rank As String * QUESTICON_LENGTH
    OutOfOrder As Byte
    
    'Maxes
    Max_CLI As Long
    
    'Main data
    CLI() As CLIRec
    Requirements As RequirementsRec

End Type

Private Type TempQuestRec

    CLI() As CLIRec

End Type

'/////////////////////////////////////////
'////////QUEST SUBS AND FUNCTIONS//
'////////////////////////////////////////
Public Function GetQuestEXP(ByVal QuestID As Long) As String

    Dim i As Long, II As Long, Count As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    With Quest(QuestID)

        For i = 1 To .Max_CLI
            For II = 1 To .CLI(i).Max_Actions

                If .CLI(i).Action(II).ActionID = ACTION_ADJUST_EXP Then
                    If .CLI(i).Action(II).MainData = vbUnchecked Then 'make sure we're adding to the player's exp and not setting it
                        Count = Count + .CLI(i).Action(II).amount
                    End If
                End If

            Next II
        Next i

    End With
    
    If Count > 0 Then
        GetQuestEXP = Format$(Count, "###,###,###,###")
    Else
        GetQuestEXP = 0
    End If
   
    ' Error Handler
    Exit Function

ErrorHandler:
    HandleError "GetQuestEXP", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Function

End Function

Public Function HasItem(ByVal ItemNum As Long) As Long

    Dim i As Long

    ' Check for subscript out of range
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If IsPlaying(MyIndex) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then

        Exit Function

    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If PlayerInv(i).num = ItemNum Then
            HasItem = HasItem + PlayerInv(i).Value
        End If

    Next
   
    ' Error Handler
    Exit Function

ErrorHandler:
    HandleError "HasItem", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Function

End Function

Public Sub MoveListItem(ByVal ListID As Byte, _
                        ByVal Index As Long, _
                        ByVal CLIIndex As Long, _
                        ByVal ArrayID As Long, _
                        ByVal Dir As Integer)

    Dim i         As Long

    Dim TempQuest As TempQuestRec

    'If the dir is -1, it means we're backtracking a slot to move the item up
    'If the dir is 1, it means we're moving ahead a slot to move the item down
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Select Case ListID

        Case LIST_CLI
            ReDim TempQuest.CLI(1 To 2)

            'Copy the slots
            TempQuest.CLI(1) = Quest(Index).CLI(ArrayID + Dir)
            TempQuest.CLI(2) = Quest(Index).CLI(ArrayID)
            
            'Paste the slots
            Quest(Index).CLI(ArrayID + Dir) = TempQuest.CLI(2)
            Quest(Index).CLI(ArrayID) = TempQuest.CLI(1)
            
        Case LIST_TASK
            ReDim TempQuest.CLI(1 To 2)
            ReDim Preserve TempQuest.CLI(1).Action(1 To 1)
            ReDim Preserve TempQuest.CLI(2).Action(1 To 1)
        
            'Copy the slots
            TempQuest.CLI(1).Action(1) = Quest(Index).CLI(CLIIndex).Action(ArrayID + Dir)
            TempQuest.CLI(2).Action(1) = Quest(Index).CLI(CLIIndex).Action(ArrayID)
            
            'Paste the slots
            Quest(Index).CLI(CLIIndex).Action(ArrayID + Dir) = TempQuest.CLI(2).Action(1)
            Quest(Index).CLI(CLIIndex).Action(ArrayID) = TempQuest.CLI(1).Action(1)

        Case Else

            Exit Sub

    End Select

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "MoveListItem", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

' //////////////////
' // Quest Editor //
' //////////////////
Public Sub QuestEditorInit()

    Dim i        As Long

    Dim SoundSet As Boolean
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    EditorIndex = frmEditor_Quest.lstIndex.ListIndex + 1
    Quest_Changed(EditorIndex) = True

    With Quest(EditorIndex)
        
        ' Basic data
        frmEditor_Quest.txtName.text = Trim$(.Name)
        frmEditor_Quest.txtDesc.text = Trim$(.Description)
        frmEditor_Quest.chkRetake.Value = .CanBeRetaken
        frmEditor_Quest.txtRank.text = Trim$(.Rank)
        frmEditor_Quest.chkUnOrder.Value = .OutOfOrder
        
        ' Gender requirement
        frmEditor_Quest.cmbGenderReq.ListIndex = .Requirements.GenderReq
        
        ' Skill requirement
        frmEditor_Quest.cmbSkillReq.ListIndex = .Requirements.SkillReq
        frmEditor_Quest.scrlSkill.Value = .Requirements.SkillLevelReq
        
        ' Basic requirements
        frmEditor_Quest.scrlAccessReq.Value = .Requirements.AccessReq
        frmEditor_Quest.scrlLevelReq.Value = .Requirements.LevelReq
        
        ' Class requirements
        frmEditor_Quest.cmbClassReq.ListIndex = .Requirements.ClassReq
        
        ' Loop for stats
        For i = 1 To Stats.Stat_Count - 1
            frmEditor_Quest.scrlStatReq(i).Value = .Requirements.Stat_Req(i)
        Next
        
        'Loop for CLI's
        frmEditor_Quest.lstTasks.Clear
        frmEditor_Quest.CLI.Clear

        For i = 1 To .Max_CLI

            If .CLI(i).ItemIndex > 0 Then
                If .CLI(i).isNPC Then frmEditor_Quest.CLI.AddItem "Meet with: " & Trim$(NPC(.CLI(i).ItemIndex).Name)
            Else

                If Not .CLI(i).isNPC Then frmEditor_Quest.CLI.AddItem "Event Interaction Only"
            End If

        Next i

    End With
    
    Call CheckStartMsg
    
    'simply set focus to the mission name textbox.
    frmEditor_Quest.txtName.SetFocus
    frmEditor_Quest.txtName.SelStart = 1
    frmEditor_Quest.txtName.SelLength = Len(frmEditor_Quest.txtName.text)

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "QuestEditorInit", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

' ///////////////////////////
' // Quest Editor CLI List //
' ///////////////////////////
Public Sub QuestEditorInitCLI()

    Dim i      As Long

    Dim Index  As Long

    Dim Tmp    As String

    Dim Indent As String
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Index = frmEditor_Quest.CLI.ListIndex + 1

    If Index < 1 Then Exit Sub

    With Quest(EditorIndex).CLI(Index)
        
        'Loop for Actions
        frmEditor_Quest.lstTasks.Clear
        
        For i = 1 To .Max_Actions

            If .Action(i).ActionID > 0 Then

                'Little more in depth because it turns code into a description easily readable by the user
                Select Case .Action(i).ActionID
                    
                    Case TASK_KILL 'gather items
                        frmEditor_Quest.lstTasks.AddItem "Kill " & .Action(i).amount & " " & Trim$(NPC(.Action(i).MainData).Name) & "('s)."
                    
                    Case TASK_GATHER 'gather items

                        If .Action(i).SecondaryData = 1 Then
                            frmEditor_Quest.lstTasks.AddItem "Gather and handover " & .Action(i).amount & " " & Trim$(Item(.Action(i).MainData).Name) & "('s)."
                        Else
                            frmEditor_Quest.lstTasks.AddItem "Gather " & .Action(i).amount & " " & Trim$(Item(.Action(i).MainData).Name) & "('s)."
                        End If
                        
                    Case TASK_VARIABLE

                        If CBool(.Action(i).MainData) = True Then
                            frmEditor_Quest.lstTasks.AddItem "Return a variable value of " & .Action(i).amount & " for (" & Variables(.Action(i).SecondaryData) & ")"
                        Else

                            If CBool(.Action(i).amount) = True Then
                                frmEditor_Quest.lstTasks.AddItem "Return a switch value of 'True' for (" & Switches(.Action(i).SecondaryData) & ")"
                            Else
                                frmEditor_Quest.lstTasks.AddItem "Return a switch value of 'False' for (" & Switches(.Action(i).SecondaryData) & ")"
                            End If
                        End If
                        
                    Case TASK_GETSKILL
                        'frmEditor_Quest.lstTasks.AddItem "Obtain level " & .Action(I).Amount & " for the " & GetSkillName(.Action(I).MainData) & " skill."
                        frmEditor_Quest.lstTasks.AddItem "SKILLS ARE NOT AN IMPLEMENTED SYSTEM IN NIN"
                        
                    Case ACTION_SETVARIABLE

                        If CBool(.Action(i).MainData) = True Then
                            frmEditor_Quest.lstTasks.AddItem "----Set player variable value to " & .Action(i).amount & " for (" & Variables(.Action(i).SecondaryData) & ")"
                        Else

                            If CBool(.Action(i).amount) = True Then
                                frmEditor_Quest.lstTasks.AddItem "----Set player switch value to 'True' for (" & Switches(.Action(i).SecondaryData) & ")"
                            Else
                                frmEditor_Quest.lstTasks.AddItem "----Set player switch value to 'False' for (" & Switches(.Action(i).SecondaryData) & ")"
                            End If
                        End If
                    
                    Case ACTION_GIVE_ITEM 'give player item(s)
                        frmEditor_Quest.lstTasks.AddItem "----Give " & .Action(i).amount & " " & Trim$(Item(.Action(i).MainData).Name) & "('s) to player."
                    
                    Case ACTION_TAKE_ITEM 'take player item(s)
                        frmEditor_Quest.lstTasks.AddItem "----Take " & .Action(i).amount & " " & Trim$(Item(.Action(i).MainData).Name) & "('s) from player."
                    
                    Case ACTION_SHOWMSG 'show the player a message

                        If .Action(i).MainData = vbChecked Then
                            frmEditor_Quest.lstTasks.AddItem "START: (" & GetColorName(.Action(i).TertiaryData) & ") Show start msg: " & """" & Trim$(.Action(i).TextHolder) & """"
                        Else

                            If Not .Action(i).QuadData = vbChecked Then
                                If .Action(i).SecondaryData = vbChecked Then
                                    frmEditor_Quest.lstTasks.AddItem "---- (" & GetColorName(.Action(i).TertiaryData) & ") Show msg if last task is incomplete: " & """" & Trim$(.Action(i).TextHolder) & """"
                                Else
                                    frmEditor_Quest.lstTasks.AddItem "---- (" & GetColorName(.Action(i).TertiaryData) & ") Show msg: " & """" & Trim$(.Action(i).TextHolder) & """"
                                End If

                            Else
                                frmEditor_Quest.lstTasks.AddItem "---- (" & GetColorName(.Action(i).TertiaryData) & ") Show msg on mission retry: " & """" & Trim$(.Action(i).TextHolder) & """"
                            End If
                        End If
                    
                    Case ACTION_ADJUST_EXP 'adjust the player's experience

                        If .Action(i).MainData = 0 Then
                            If .Action(i).amount > 0 Then Tmp = "+"
                            frmEditor_Quest.lstTasks.AddItem "----Modify Player EXP by " & Tmp & .Action(i).amount
                        Else
                            frmEditor_Quest.lstTasks.AddItem "----Set Player EXP to " & .Action(i).amount
                        End If
                    
                    Case ACTION_ADJUST_LVL 'adjust the player's level

                        If .Action(i).MainData = 0 Then
                            If .Action(i).amount > 0 Then Tmp = "+"
                            frmEditor_Quest.lstTasks.AddItem "---- Modify Player Level by " & Tmp & .Action(i).amount
                        Else
                            frmEditor_Quest.lstTasks.AddItem "---- Set Player Level to " & .Action(i).amount
                        End If
                    
                    Case ACTION_ADJUST_STAT_LVL 'adjust the player's stat level

                        If .Action(i).MainData = 0 Then
                            If .Action(i).amount > 0 Then Tmp = "+"
                            frmEditor_Quest.lstTasks.AddItem "---- Modify Player's " & GetStatName(.Action(i).SecondaryData) & " Level by " & Tmp & .Action(i).amount
                        Else
                            frmEditor_Quest.lstTasks.AddItem "---- Set Player's " & GetStatName(.Action(i).SecondaryData) & " Level to " & .Action(i).amount
                        End If
                        
                    Case ACTION_ADJUST_SKILL_LVL 'adjust the player's skill level

                        If .Action(i).MainData = 0 Then
                            If .Action(i).amount > 0 Then Tmp = "+"
                            'frmEditor_Quest.lstTasks.AddItem "---- Modify Player's " & GetSkillName(.Action(I).SecondaryData) & " level by " & Tmp & .Action(I).Amount
                        Else
                            'frmEditor_Quest.lstTasks.AddItem "---- Set Player's " & GetSkillName(.Action(I).SecondaryData) & " level to " & .Action(I).Amount
                        End If

                        frmEditor_Quest.lstTasks.AddItem "SKILLS ARE NOT AN IMPLEMENTED SYSTEM IN NIN"
                        
                    Case ACTION_ADJUST_SKILL_EXP 'adjust the player's skill exp

                        If .Action(i).MainData = 0 Then
                            If .Action(i).amount > 0 Then Tmp = "+"
                            'frmEditor_Quest.lstTasks.AddItem "---- Modify Player's " & GetSkillName(.Action(I).SecondaryData) & " EXP by " & Tmp & .Action(I).Amount
                        Else
                            'frmEditor_Quest.lstTasks.AddItem "---- Set Player's " & GetSkillName(.Action(I).SecondaryData) & " EXP to " & .Action(I).Amount
                        End If

                        frmEditor_Quest.lstTasks.AddItem "SKILLS ARE NOT AN IMPLEMENTED SYSTEM IN NIN"
                        
                    Case ACTION_ADJUST_STAT_POINTS 'adjust the player's stat points

                        If .Action(i).MainData = 0 Then
                            If .Action(i).amount > 0 Then Tmp = "+"
                            frmEditor_Quest.lstTasks.AddItem "---- Modify Player's stat points by " & Tmp & .Action(i).amount
                        Else
                            frmEditor_Quest.lstTasks.AddItem "---- Set Player's stat points to " & .Action(i).amount
                        End If
                        
                    Case ACTION_WARP
                        frmEditor_Quest.lstTasks.AddItem "---- Warp player to Map: " & .Action(i).amount & " (X" & .Action(i).MainData & ", Y" & .Action(i).SecondaryData & ")"
                    
                    Case ACTION_PLAYSOUND

                        If .Action(i).SecondaryData = 0 Then 'player
                            frmEditor_Quest.lstTasks.AddItem "---- Play Sound [to Player]: " & SoundCache(.Action(i).MainData)
                        ElseIf .Action(i).SecondaryData = 0 Then 'map
                            frmEditor_Quest.lstTasks.AddItem "---- Play Sound [to Map]: " & SoundCache(.Action(i).MainData)
                        ElseIf .Action(i).SecondaryData = 0 Then 'all
                            frmEditor_Quest.lstTasks.AddItem "---- Play Sound [to Everyone]: " & SoundCache(.Action(i).MainData)
                        End If
                        
                    Case Else

                        Exit Sub
                
                End Select

            End If
            
            Tmp = vbNullString
        Next i

    End With
    
    Call CheckStartMsg

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "QuestEditorInitCLI", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Public Sub DeleteAction(ByVal QuestID As Long, _
                        ByVal Index As Long, _
                        ByVal ActionID As Long)

    Dim i As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    With Quest(QuestID).CLI(Index)

        For i = 1 To .Max_Actions

            If i >= ActionID Then
                Call ZeroMemory(ByVal VarPtr(.Action(i)), LenB(.Action(i)))
                
                'start swaping the following slots
                If i + 1 <= .Max_Actions Then
                    .Action(i) = .Action(i + 1)
                End If
            End If

        Next i
        
        Call ZeroMemory(ByVal VarPtr(.Action(.Max_Actions)), LenB(.Action(.Max_Actions)))
        .Max_Actions = .Max_Actions - 1

        If .Max_Actions > 0 Then
            ReDim Preserve .Action(1 To .Max_Actions)
        Else
            ReDim .Action(0 To 0)
        End If

    End With

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "DeleteAction", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Public Sub DeleteCLI(ByVal QuestID As Long, ByVal Index As Long)

    Dim i As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    With Quest(QuestID)

        For i = 1 To .Max_CLI

            If i >= Index Then
                Call ZeroMemory(ByVal VarPtr(.CLI(i)), LenB(.CLI(i)))
                
                'start swaping the following slots
                If i + 1 <= .Max_CLI Then
                    .CLI(i) = .CLI(i + 1)
                End If
            End If

        Next i
        
        Call ZeroMemory(ByVal VarPtr(.CLI(.Max_CLI)), LenB(.CLI(.Max_CLI)))
        .Max_CLI = .Max_CLI - 1

        If .Max_CLI > 0 Then
            ReDim Preserve .CLI(1 To .Max_CLI)
        Else
            ReDim .CLI(0 To 0)
        End If

    End With
    
    Call QuestEditorInit

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "DeleteCLI", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Public Sub CheckStartMsg()

    Dim i As Long, II As Long, III As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    For i = 1 To MAX_QUESTS
        For II = 1 To Quest(i).Max_CLI
            For III = 1 To Quest(i).CLI(II).Max_Actions

                If Quest(i).CLI(II).Action(III).ActionID = ACTION_SHOWMSG Then
                    If Quest(i).CLI(II).Action(III).MainData = vbChecked Then
                        frmEditor_Quest.chkStart.Value = vbUnchecked

                        Exit Sub

                    End If
                End If

            Next III
        Next II
    Next i
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "CheckStartMsg", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Public Sub CheckResponseMsg(ByVal QuestID As Long, _
                            ByVal CLIIndex As Long, _
                            ByVal SearchStopIndex As Long)

    Dim i As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    frmEditor_Quest.chkRes.Enabled = False

    If QuestID < 1 Or QuestID > MAX_QUESTS Then Exit Sub
    If CLIIndex < 1 Then Exit Sub
    
    For i = 1 To SearchStopIndex

        If Quest(QuestID).CLI(CLIIndex).Action(i).ActionID > 0 And Quest(QuestID).CLI(CLIIndex).Action(i).ActionID < 5 Then 'find a task
            frmEditor_Quest.chkRes.Enabled = True 'found a task, allow access to the checkbox

            Exit Sub

        End If

    Next i
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "CheckResponseMsg", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Public Function IsNPCInAnotherQuest(ByVal NPCIndex As Long, CurQuest As Long) As Boolean

    Dim i As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    IsNPCInAnotherQuest = False
    
    For i = 1 To MAX_QUESTS

        If i <> CurQuest Then
            If Quest(i).Max_CLI > 0 Then 'prevent subscript out of range within the following prcedure
                If Quest(i).CLI(1).ItemIndex = NPCIndex Then
                    IsNPCInAnotherQuest = True

                    Exit Function

                End If
            End If
        End If

    Next i
   
    ' Error Handler
    Exit Function

ErrorHandler:
    HandleError "IsNPCInAnotherQuest", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Function

End Function

Public Function DoesNPCStartQuest(ByVal npcNum As Long) As Long

    Dim i As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    DoesNPCStartQuest = 0
    
    For i = 1 To MAX_QUESTS

        If Quest(i).Max_CLI > 0 Then
            If Quest(i).CLI(1).ItemIndex = npcNum Then
                DoesNPCStartQuest = i 'return the quest number

                Exit Function

            End If
        End If

    Next i
   
    ' Error Handler
    Exit Function

ErrorHandler:
    HandleError "DoesNPCStartQuest", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Function

End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~DATA HANDLER~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub HandleUpdateQuest(ByVal Index As Long, _
                             ByRef data() As Byte, _
                             ByVal StartAddr As Long, _
                             ByVal ExtraVar As Long)

    Dim buffer   As clsBuffer

    Dim i        As Long, II As Long

    Dim QuestNum As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    QuestNum = buffer.ReadLong

    With Quest(QuestNum)
        
        .Name = buffer.ReadString
        .Description = buffer.ReadString
        .CanBeRetaken = buffer.ReadLong
        .Max_CLI = buffer.ReadLong
        .Rank = buffer.ReadString
        .OutOfOrder = buffer.ReadByte
            
        .Requirements.AccessReq = buffer.ReadLong
        .Requirements.ClassReq = buffer.ReadLong
        .Requirements.GenderReq = buffer.ReadLong
        .Requirements.LevelReq = buffer.ReadLong
        .Requirements.SkillLevelReq = buffer.ReadLong
        .Requirements.SkillReq = buffer.ReadLong
            
        For i = 1 To Stats.Stat_Count - 1
            .Requirements.Stat_Req(i) = buffer.ReadLong
        Next i
            
        If .Max_CLI > 0 Then
            ReDim .CLI(1 To .Max_CLI)
            
            For i = 1 To .Max_CLI
                .CLI(i).ItemIndex = buffer.ReadLong
                .CLI(i).isNPC = buffer.ReadLong
                .CLI(i).Max_Actions = buffer.ReadLong
                    
                If .CLI(i).Max_Actions > 0 Then
                    ReDim Preserve .CLI(i).Action(1 To .CLI(i).Max_Actions)
                    
                    For II = 1 To .CLI(i).Max_Actions
                        .CLI(i).Action(II).TextHolder = buffer.ReadString
                        .CLI(i).Action(II).ActionID = buffer.ReadLong
                        .CLI(i).Action(II).amount = buffer.ReadLong
                        .CLI(i).Action(II).MainData = buffer.ReadLong
                        .CLI(i).Action(II).QuadData = buffer.ReadLong
                        .CLI(i).Action(II).SecondaryData = buffer.ReadLong
                        .CLI(i).Action(II).TertiaryData = buffer.ReadLong
                    Next II

                End If

            Next i

        End If

    End With
    
    Set buffer = Nothing
    
    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "HandleUpdateQuest", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Public Sub HandleQuestEditor()

    Dim i As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    With frmEditor_Quest
        Editor = EDITOR_QUESTS
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_QUESTS
            .lstIndex.AddItem i & ": " & Trim$(Quest(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        QuestEditorInit
    End With

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "HandleQuestEditor", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Public Sub HandleQuestRequest(ByVal Index As Long, _
                              ByRef data() As Byte, _
                              ByVal StartAddr As Long, _
                              ByVal ExtraVar As Long)

    Dim buffer   As clsBuffer

    Dim QuestNum As Long

    Dim Msg      As String, Color As Byte
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    QuestNum = buffer.ReadLong
    Set buffer = Nothing
    
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Sub
    QuestRequest = QuestNum
    
    If Quest(QuestRequest).Max_CLI > 0 Then
        Msg = Trim$(Quest(QuestRequest).CLI(1).Action(1).TextHolder)
        Color = Quest(QuestRequest).CLI(1).Action(1).TertiaryData
        Call AddText(Trim$(Quest(QuestRequest).Name) & ": " & Msg, Color)
    End If
    
    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "HandleQuestRequest", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~MODIFIERS~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub ClearQuest(ByVal Index As Long)
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Quest(Index)), LenB(Quest(Index)))
    Quest(Index).Name = vbNullString
    Quest(Index).Description = vbNullString
    Quest(Index).Max_CLI = 0
    Quest(Index).CanBeRetaken = 0

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "ClearQuest", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Sub ClearQuests()

    Dim i As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "ClearQuests", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Public Sub QuestEditorSave()

    Dim i As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    For i = 1 To MAX_QUESTS

        If Quest_Changed(i) Then
            Call SendSaveQuest(i)
        End If

    Next
    
    Unload frmEditor_Quest
    Editor = 0
    ClearChanged_Quest

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "QuestEditorSave", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~TCP~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Sub SendSaveQuest(ByVal QuestNum As Long)

    Dim buffer As clsBuffer

    Dim i      As Long, II As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    
    buffer.WriteLong CSaveQuest
        
    With Quest(QuestNum)
        buffer.WriteLong QuestNum
        buffer.WriteString .Name
        buffer.WriteString .Description
        buffer.WriteLong .CanBeRetaken
        buffer.WriteLong .Max_CLI
        buffer.WriteString .Rank
        buffer.WriteByte .OutOfOrder
            
        buffer.WriteLong .Requirements.AccessReq
        buffer.WriteLong .Requirements.ClassReq
        buffer.WriteLong .Requirements.GenderReq
        buffer.WriteLong .Requirements.LevelReq
        buffer.WriteLong .Requirements.SkillLevelReq
        buffer.WriteLong .Requirements.SkillReq
            
        For i = 1 To Stats.Stat_Count - 1
            buffer.WriteLong .Requirements.Stat_Req(i)
        Next i
            
        If .Max_CLI > 0 Then

            For i = 1 To .Max_CLI
                buffer.WriteLong .CLI(i).ItemIndex
                buffer.WriteLong .CLI(i).isNPC
                buffer.WriteLong .CLI(i).Max_Actions
                    
                For II = 1 To .CLI(i).Max_Actions
                    buffer.WriteString .CLI(i).Action(II).TextHolder
                    buffer.WriteLong .CLI(i).Action(II).ActionID
                    buffer.WriteLong .CLI(i).Action(II).amount
                    buffer.WriteLong .CLI(i).Action(II).MainData
                    buffer.WriteLong .CLI(i).Action(II).QuadData
                    buffer.WriteLong .CLI(i).Action(II).SecondaryData
                    buffer.WriteLong .CLI(i).Action(II).TertiaryData
                Next II
            Next i

        End If
            
    End With
        
    Call SendData(buffer.ToArray())
        
    Set buffer = Nothing
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "SendSaveQuest", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Public Sub HandleSendMusic(ByVal Index As Long, _
                           ByRef data() As Byte, _
                           ByVal StartAddr As Long, _
                           ByVal ExtraVar As Long)

    Dim buffer As clsBuffer, musicName As String
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    musicName = buffer.ReadString
    
    Call Audio.PlayMusic(musicName)
    Set buffer = Nothing
    
    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "HandleSendMusic", "modQuest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub


