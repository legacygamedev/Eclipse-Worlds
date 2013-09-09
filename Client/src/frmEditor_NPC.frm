VERSION 5.00
Begin VB.Form frmEditor_NPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NPC Editor"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_NPC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   606
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   567
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   6360
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   360
      Width           =   600
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   28
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   30
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   29
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   8415
      Left            =   3360
      TabIndex        =   32
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txtMP 
         Height          =   285
         Left            =   3360
         TabIndex        =   12
         Text            =   "0"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.ComboBox cmbMusic 
         Height          =   300
         Left            =   680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2040
         Width           =   1815
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CheckBox chkFactionThreat 
         Caption         =   "Faction Threat"
         Height          =   255
         Left            =   3000
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   "Other faction members will defend this NPC"
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtAttackSay 
         Height          =   285
         Left            =   3360
         TabIndex        =   10
         Top             =   2760
         Width           =   1575
      End
      Begin VB.ComboBox cmbFaction 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":038A
         Left            =   1080
         List            =   "frmEditor_NPC.frx":0397
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtTitle 
         Height          =   270
         Left            =   1080
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.HScrollBar scrlLevel 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.HScrollBar scrlDamage 
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtSpawnSecs 
         Height          =   285
         Left            =   3600
         TabIndex        =   14
         Text            =   "0"
         ToolTipText     =   "In seconds."
         Top             =   3480
         Width           =   1335
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   2040
         TabIndex        =   16
         Top             =   4200
         Width           =   1215
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cmbBehavior 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":03AD
         Left            =   1080
         List            =   "frmEditor_NPC.frx":03BD
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2760
         Width           =   1815
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Text            =   "0"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txtEXP 
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Text            =   "0"
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Frame fraStats 
         Caption         =   "Stats"
         Height          =   1455
         Left            =   120
         TabIndex        =   33
         Top             =   4920
         Width           =   4815
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   18
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   3
            Left            =   3240
            TabIndex        =   19
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   5
            Left            =   1680
            TabIndex        =   21
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Str: 0"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "End: 0"
            Height          =   180
            Index           =   2
            Left            =   1680
            TabIndex        =   37
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Int: 0"
            Height          =   180
            Index           =   3
            Left            =   3240
            TabIndex        =   36
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Agi: 0"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   1425
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Spi: 0"
            Height          =   180
            Index           =   5
            Left            =   1680
            TabIndex        =   34
            Top             =   840
            Width           =   450
         End
      End
      Begin VB.Frame fraDrop 
         Caption         =   "Drop: 1"
         Height          =   1815
         Left            =   120
         TabIndex        =   48
         Top             =   6360
         Width           =   4815
         Begin VB.TextBox txtChance 
            Height          =   285
            Left            =   2880
            TabIndex        =   25
            Text            =   "0"
            ToolTipText     =   "Use 0, 1, number%, 1/number, or decimal values."
            Top             =   720
            Width           =   1815
         End
         Begin VB.HScrollBar scrlNum 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   26
            Top             =   1080
            Width           =   3495
         End
         Begin VB.HScrollBar scrlValue 
            Height          =   255
            Left            =   1200
            TabIndex        =   27
            Top             =   1440
            Width           =   3495
         End
         Begin VB.HScrollBar scrlDrop 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   24
            Top             =   240
            Value           =   1
            Width           =   4575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Chance:"
            Height          =   180
            Left            =   2160
            TabIndex        =   52
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   630
         End
         Begin VB.Label lblNum 
            AutoSize        =   -1  'True
            Caption         =   "Number: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   51
            Top             =   1080
            Width           =   795
         End
         Begin VB.Label lblItemName 
            AutoSize        =   -1  'True
            Caption         =   "Item: None"
            Height          =   180
            Left            =   120
            TabIndex        =   50
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            Caption         =   "Value: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   49
            Top             =   1440
            UseMnemonic     =   0   'False
            Width           =   645
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   4800
            Y1              =   600
            Y2              =   600
         End
      End
      Begin VB.CheckBox chkSwapVisibility 
         Caption         =   "Show Spells"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Frame fraSpell 
         Caption         =   "Spell: 1"
         Height          =   1455
         Left            =   120
         TabIndex        =   62
         Top             =   4920
         Visible         =   0   'False
         Width           =   4815
         Begin VB.HScrollBar scrlSpellNum 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   23
            Top             =   1080
            Width           =   3495
         End
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   22
            Top             =   240
            Value           =   1
            Width           =   4575
         End
         Begin VB.Label lblSpellName 
            AutoSize        =   -1  'True
            Caption         =   "Spell: None"
            Height          =   180
            Left            =   120
            TabIndex        =   64
            Top             =   720
            Width           =   870
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   120
            X2              =   4920
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblSpellNum 
            AutoSize        =   -1  'True
            Caption         =   "Number: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   63
            Top             =   1080
            Width           =   795
         End
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "MP:"
         Height          =   180
         Left            =   3000
         TabIndex        =   61
         Top             =   3120
         Width           =   300
      End
      Begin VB.Label Label6 
         Caption         =   "Music:"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   2520
         TabIndex        =   59
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label lblAttackSay 
         AutoSize        =   -1  'True
         Caption         =   "Say:"
         Height          =   180
         Left            =   3000
         TabIndex        =   57
         Top             =   2760
         UseMnemonic     =   0   'False
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Faction:"
         Height          =   180
         Left            =   120
         TabIndex        =   56
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   180
         Left            =   120
         TabIndex        =   55
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   405
      End
      Begin VB.Label lblLevel 
         Caption         =   "Level: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblDamage 
         Caption         =   "Damage: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Spawn:"
         Height          =   180
         Left            =   3000
         TabIndex        =   47
         Top             =   3480
         UseMnemonic     =   0   'False
         Width           =   540
      End
      Begin VB.Label lblAnimation 
         Caption         =   "Animation: None"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   44
         Top             =   1320
         Width           =   660
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Behavior:"
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   2760
         UseMnemonic     =   0   'False
         Width           =   720
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   41
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Experience:"
         Height          =   180
         Left            =   120
         TabIndex        =   40
         Top             =   3480
         Width           =   900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "HP:"
         Height          =   180
         Left            =   120
         TabIndex        =   39
         Top             =   3120
         Width           =   285
      End
   End
   Begin VB.Frame fraNPC 
      Caption         =   "NPC List"
      Height          =   8895
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   315
         Left            =   2400
         TabIndex        =   68
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   315
         Left            =   1680
         TabIndex        =   66
         Top             =   240
         Width           =   615
      End
      Begin VB.ListBox lstIndex 
         Height          =   7860
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_NPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DropIndex As Long
Private SpellIndex As Long
Private TmpIndex As Long

Private Sub chkFactionThreat_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If chkFactionThreat.Value = 1 Then
        NPC(EditorIndex).FactionThreat = True
    Else
        NPC(EditorIndex).FactionThreat = False
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "chkFactionThreat_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkSwapVisibility_Click()
    fraSpell.Visible = Not fraSpell.Visible
    fraStats.Visible = Not fraStats.Visible
End Sub

Private Sub cmbBehavior_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    NPC(EditorIndex).Behavior = cmbBehavior.ListIndex
    
    If NPC(EditorIndex).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or NPC(EditorIndex).Behavior = NPC_BEHAVIOR_GUARD Then
        frmEditor_NPC.txtAttackSay.Enabled = True
        frmEditor_NPC.lblAttackSay.Enabled = True
    Else
        frmEditor_NPC.txtAttackSay.Enabled = False
        frmEditor_NPC.lblAttackSay.Enabled = False
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmbBehavior_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbFaction_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    NPC(EditorIndex).Faction = cmbFaction.ListIndex
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmbFaction_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbMusic_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbMusic.ListIndex > 0 Then
        NPC(EditorIndex).Music = cmbMusic.List(cmbMusic.ListIndex)
    Else
        NPC(EditorIndex).Music = vbNullString
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdMusic_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex > 0 Then
        NPC(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        NPC(EditorIndex).Sound = vbNullString
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearNPC EditorIndex
    
    TmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    NPCEditorInit
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub


Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    frmMain.SubDaFocus Me.hWnd
    scrlSprite.max = NumCharacters
    scrlAnimation.max = MAX_ANIMATIONS
    scrlDrop.max = MAX_NPC_DROPS
    scrlLevel.max = MAX_LEVEL
    txtName.MaxLength = NAME_LENGTH
    txtSearch.MaxLength = NAME_LENGTH
    txtTitle.MaxLength = NAME_LENGTH
    scrlNum.max = MAX_ITEMS
    scrlSpell.max = MAX_NPC_SPELLS
    scrlSpellNum.max = MAX_SPELLS
    
    ' Resize the sprite pictures
    If NumCharacters > 0 Then
        frmEditor_NPC.picSprite.Height = Tex_Character(1).Height / 4
        frmEditor_NPC.picSprite.Width = Tex_Character(1).Width / 4
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "Form_Load", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdSave_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    EditorSave = True
    Call NPCEditorSave
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCancel_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Unload frmEditor_NPC
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    frmMain.UnsubDaFocus Me.hWnd
    If EditorSave = False Then
        Call NPCEditorCancel
    Else
        EditorSave = False
    End If
    frmAdmin.chkEditor(EDITOR_NPC).Value = False
    BringWindowToTop (frmAdmin.hWnd)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "Form_Unload", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstIndex_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    NPCEditorInit
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlAnimation_Change()
    Dim sString As String
    
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlAnimation.Value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.Value).name)
    lblAnimation.Caption = "Animation: " & sString
    NPC(EditorIndex).Animation = scrlAnimation.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlDamage_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblDamage.Caption = "Damage: " & scrlDamage.Value
    NPC(EditorIndex).Damage = scrlDamage.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlDamage_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlLevel_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblLevel.Caption = "Level: " & scrlLevel.Value
    NPC(EditorIndex).Level = scrlLevel.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlLevel_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlSpell_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    SpellIndex = scrlSpell.Value
    fraSpell.Caption = "Spell: " & SpellIndex
    scrlSpellNum.Value = NPC(EditorIndex).Spell(SpellIndex)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlSpell_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlSpellNum_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblSpellNum.Caption = "Number: " & scrlSpellNum.Value

    If scrlSpellNum.Value > 0 Then
        lblSpellName.Caption = "Spell: " & Trim$(Spell(scrlSpellNum.Value).name)
    Else
        lblSpellName.Caption = "Spell: None"
    End If
    NPC(EditorIndex).Spell(SpellIndex) = scrlSpellNum.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlSpellNum_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlSprite_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblSprite.Caption = "Sprite: " & scrlSprite.Value
    NPC(EditorIndex).Sprite = scrlSprite.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlDrop_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    DropIndex = scrlDrop.Value
    fraDrop.Caption = "Drop: " & DropIndex
    txtChance.text = NPC(EditorIndex).DropChance(DropIndex)
    scrlNum.Value = NPC(EditorIndex).DropItem(DropIndex)
    scrlValue.Value = NPC(EditorIndex).DropValue(DropIndex)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlDrop_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlRange_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblRange.Caption = "Range: " & scrlRange.Value
    NPC(EditorIndex).Range = scrlRange.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlNum_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblNum.Caption = "Number: " & scrlNum.Value

    If scrlNum.Value > 0 Then
        lblItemName.Caption = "Item: " & Trim$(item(scrlNum.Value).name)
    Else
        lblItemName.Caption = "Item: None"
    End If
    NPC(EditorIndex).DropItem(DropIndex) = scrlNum.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlNum_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlStat_Change(Index As Integer)
    Dim prefix As String
    
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            prefix = "Str: "
        Case 2
            prefix = "End: "
        Case 3
            prefix = "Int: "
        Case 4
            prefix = "Agi: "
        Case 5
            prefix = "Spi: "
    End Select
    lblStat(Index).Caption = prefix & scrlStat(Index).Value
    NPC(EditorIndex).Stat(Index) = scrlStat(Index).Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlStat_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlValue_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblValue.Caption = "Value: " & scrlValue.Value
    NPC(EditorIndex).DropValue(DropIndex) = scrlValue.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlValue_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtAttackSay_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    NPC(EditorIndex).AttackSay = txtAttackSay.text
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtAttackSay_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtExp_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not IsNumeric(txtExp.text) Then txtExp.text = 0
    If txtExp.text > MAX_LONG Then txtExp.text = MAX_LONG
    If txtExp.text < 0 Then txtExp.text = 0
    NPC(EditorIndex).Exp = txtExp.text
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtExp_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtHP_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not IsNumeric(txtHP.text) Then txtHP.text = 0
    If txtHP.text > MAX_LONG Then txtHP.text = MAX_LONG
    If txtHP.text < 0 Then txtHP.text = 0
    NPC(EditorIndex).HP = txtHP.text
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtHP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtMP_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not IsNumeric(txtMP.text) Then txtMP.text = 0
    If txtMP.text > MAX_LONG Then txtMP.text = MAX_LONG
    If txtMP.text < 0 Then txtMP.text = 0
    NPC(EditorIndex).MP = txtMP.text
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtMP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    TmpIndex = lstIndex.ListIndex
    NPC(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSpawnSecs_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not IsNumeric(txtSpawnSecs.text) Then txtSpawnSecs.text = 0
    If txtSpawnSecs.text > MAX_LONG Then txtSpawnSecs.text = MAX_LONG
    If txtSpawnSecs.text < 0 Then txtSpawnSecs.text = 0
    NPC(EditorIndex).SpawnSecs = txtSpawnSecs.text
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtSpawnSecs_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtChance_Validate(Cancel As Boolean)
    Dim i() As String
    
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not IsNumeric(txtChance.text) And Not Right$(txtChance.text, 1) = "%" And Not InStr(1, txtChance.text, "/") > 0 And Not InStr(1, txtChance.text, ".") Then
        txtChance.text = "0"
        NPC(EditorIndex).DropChance(DropIndex) = 0
        Exit Sub
    End If
    
    If Right$(txtChance.text, 1) = "%" Then
        txtChance.text = Left$(txtChance.text, Len(txtChance.text) - 1) / 100
    ElseIf InStr(1, txtChance.text, "/") > 0 Then
        i = Split(txtChance.text, "/")
        txtChance.text = Int(i(0) / i(1) * 1000) / 1000
    End If
    
    If txtChance.text > 1 Then
        txtChance.text = "1"
    ElseIf txtChance.text < 0 Then
        txtChance.text = "0"
    End If
    
    NPC(EditorIndex).DropChance(DropIndex) = txtChance.text
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtChance_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtTitle_Validate(Cancel As Boolean)
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    NPC(EditorIndex).title = Trim$(txtTitle.text)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtTitle_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtName.SelStart = Len(txtName)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtName_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtTitle_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtTitle.SelStart = Len(txtTitle)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtTitle_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtAttackSay_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtAttackSay.SelStart = Len(txtAttackSay)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtAttackSay_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtHP_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtHP.SelStart = Len(txtHP)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtHP_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtMP_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtMP.SelStart = Len(txtMP)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtMP_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSpawnSecs_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtSpawnSecs.SelStart = Len(txtSpawnSecs)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtSpawnSecs_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtEXP_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtExp.SelStart = Len(txtExp)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtEXP_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtChance_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtChance.SelStart = Len(txtChance)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtChance_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_Change()
    Dim Find As String, i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 0 To lstIndex.ListCount - 1
        Find = Trim$(i + 1 & ": " & txtSearch.text)
        
        ' Make sure we dont try to check a name that's too small
        If Len(lstIndex.List(i)) >= Len(Find) Then
            If UCase$(Mid$(Trim$(lstIndex.List(i)), 1, Len(Find))) = UCase$(Find) Then
                lstIndex.ListIndex = i
                Exit For
            End If
        End If
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtSearch_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtSearch.SelStart = Len(txtSearch)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtSearch_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If KeyAscii = vbKeyReturn Then
        cmdSave_Click
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyEscape Then
        cmdCancel_Click
        KeyAscii = 0
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "Form_KeyPress", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCopy_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    TmpIndex = lstIndex.ListIndex
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdCopy_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdPaste_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lstIndex.RemoveItem EditorIndex - 1
    Call CopyMemory(ByVal VarPtr(NPC(EditorIndex)), ByVal VarPtr(NPC(TmpIndex + 1)), LenB(NPC(TmpIndex + 1)))
    lstIndex.AddItem EditorIndex & ": " & Trim$(NPC(EditorIndex).name), EditorIndex - 1
    lstIndex.ListIndex = EditorIndex - 1
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdPaste_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
