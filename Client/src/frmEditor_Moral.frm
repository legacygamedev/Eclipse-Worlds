VERSION 5.00
Begin VB.Form frmEditor_Moral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Moral Editor"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5865
   Icon            =   "frmEditor_Moral.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   271
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   391
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPlayerBlocked 
      Caption         =   "Player Blocked"
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   3495
      Left            =   2760
      TabIndex        =   7
      Top             =   0
      Width           =   3015
      Begin VB.HScrollBar scrlColor 
         Height          =   255
         Left            =   120
         Max             =   17
         TabIndex        =   2
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2775
      End
      Begin VB.CheckBox chkCanPK 
         Caption         =   "Can PvP"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1560
         Width           =   975
      End
      Begin VB.CheckBox chkCanCast 
         Caption         =   "Can Cast"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox chkCanUseItem 
         Caption         =   "Can Use Item"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox chkDropItems 
         Caption         =   "Drop Items On Death"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CheckBox chkLoseExp 
         Caption         =   "Lose Experience On Death"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3000
         Width           =   2295
      End
      Begin VB.CheckBox chkCanDropItem 
         Caption         =   "Can Drop Item"
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox chkCanPickupItem 
         Caption         =   "Can Pickup Item"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color: Black"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   3600
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Moral List"
      Height          =   3975
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   315
         Left            =   1800
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   315
         Left            =   1080
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.ListBox lstIndex 
         Height          =   2985
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmEditor_Moral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TmpIndex As Long

Private Sub chkCanPickupItem_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_MORALS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Moral(EditorIndex).CanPickupItem = chkCanPickupItem.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "chkCanPickupItem_Click", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkCanCast_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_MORALS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Moral(EditorIndex).CanCast = chkCanCast.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "chkCanCast_Click", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkCanDropItem_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_MORALS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Moral(EditorIndex).CanDropItem = chkCanDropItem.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "chkCanDropItem_Click", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkCanPK_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_MORALS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Moral(EditorIndex).CanPK = chkCanPK.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "chkCanPK_Click", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkCanUseItem_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_MORALS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Moral(EditorIndex).CanUseItem = chkCanUseItem.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "chkCanUseItem_Click", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkDropItems_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_MORALS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Moral(EditorIndex).DropItems = chkDropItems.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "chkDropItems_Click", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkLoseExp_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_MORALS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Moral(EditorIndex).LoseExp = chkLoseExp.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "chkLoseExp_Click", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkPlayerBlocked_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_MORALS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Moral(EditorIndex).PlayerBlocked = chkPlayerBlocked.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "chkPlayerBlocked_Click", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_MORALS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearMoral EditorIndex
    
    TmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Moral(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    
    MoralEditorInit
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdSave_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_MORALS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    EditorSave = True
    MoralEditorSave
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCancel_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_MORALS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Unload frmEditor_Moral
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorSave = False Then
        MoralEditorCancel
    Else
        EditorSave = False
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "Form_Unload", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstIndex_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_MORALS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    MoralEditorInit
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "1stIndex_Click", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtName.MaxLength = NAME_LENGTH
    txtSearch.MaxLength = NAME_LENGTH
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "Form_Load", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlColor_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblColor.Caption = "Color: " & GetColorName(scrlColor.Value)
    Moral(EditorIndex).Color = scrlColor.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlColor_Change", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim TmpIndex As Long

    If EditorIndex < 1 Or EditorIndex > MAX_MORALS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    TmpIndex = lstIndex.ListIndex
    Moral(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Moral(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtName.SelStart = Len(txtName)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtName_GotFocus", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    HandleError "txtSearch_Change", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtSearch.SelStart = Len(txtSearch)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtSearch_GotFocus", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    HandleError "Form_KeyPress", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCopy_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    TmpIndex = lstIndex.ListIndex
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdCopy_Click", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdPaste_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
        
    lstIndex.RemoveItem EditorIndex - 1
    Call CopyMemory(ByVal VarPtr(Moral(EditorIndex)), ByVal VarPtr(Moral(TmpIndex + 1)), LenB(Moral(TmpIndex + 1)))
    lstIndex.AddItem EditorIndex & ": " & Trim$(Moral(EditorIndex).name), EditorIndex - 1
    lstIndex.ListIndex = EditorIndex - 1
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdPaste_Click", "frmEditor_Moral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
