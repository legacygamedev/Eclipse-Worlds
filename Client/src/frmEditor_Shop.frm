VERSION 5.00
Begin VB.Form frmEditor_Shop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop Editor"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Shop.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   335
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   4455
      Left            =   3360
      TabIndex        =   17
      Top             =   0
      Width           =   5295
      Begin VB.HScrollBar scrlSell 
         Height          =   255
         LargeChange     =   100
         Left            =   2640
         Max             =   1000
         Min             =   1
         TabIndex        =   3
         Top             =   1080
         Value           =   100
         Width           =   2505
      End
      Begin VB.CheckBox chkCanFix 
         Caption         =   "Can Fix"
         Height          =   180
         Left            =   4200
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtCostValue2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   9
         Text            =   "1"
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox txtCostValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   7
         Text            =   "1"
         Top             =   1800
         Width           =   615
      End
      Begin VB.ComboBox cmbCostItem 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtItemValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   5
         Text            =   "1"
         Top             =   1440
         Width           =   615
      End
      Begin VB.ComboBox cmbItem 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1440
         Width           =   3015
      End
      Begin VB.ListBox lstTradeItem 
         Height          =   1230
         ItemData        =   "frmEditor_Shop.frx":038A
         Left            =   120
         List            =   "frmEditor_Shop.frx":03A6
         TabIndex        =   12
         Top             =   2880
         Width           =   5055
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   4455
      End
      Begin VB.HScrollBar scrlBuy 
         Height          =   255
         LargeChange     =   100
         Left            =   120
         Max             =   1000
         Min             =   1
         TabIndex        =   2
         Top             =   1080
         Value           =   100
         Width           =   2460
      End
      Begin VB.CommandButton cmdDeleteTrade 
         Caption         =   "Delete"
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   2520
         Width           =   2415
      End
      Begin VB.ComboBox cmbCostItem2 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label lblSell 
         Caption         =   "Sell Rate: 100%"
         Height          =   255
         Left            =   2640
         TabIndex        =   27
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Price 2:"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Left            =   3960
         TabIndex        =   24
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Left            =   3960
         TabIndex        =   23
         Top             =   1800
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Left            =   3960
         TabIndex        =   21
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   480
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblBuy 
         AutoSize        =   -1  'True
         Caption         =   "Buy Rate: 100%"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   2460
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Shop List"
      Height          =   4935
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   315
         Left            =   2400
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   315
         Left            =   1680
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
      Begin VB.ListBox lstIndex 
         Height          =   4155
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   6960
      TabIndex        =   15
      Top             =   4560
      Width           =   1575
   End
End
Attribute VB_Name = "frmEditor_Shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TmpIndex As Long

Private Sub chkCanFix_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Shop(EditorIndex).CanFix = chkCanFix.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "chkCanFix_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearShop EditorIndex
    
    TmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Shop(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    
    ShopEditorInit
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdSave_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    EditorSave = True
    Call ShopEditorSave
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCancel_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Unload frmEditor_Shop
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdUpdate_Click()
    Dim Index As Long
    Dim tmpPos As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    tmpPos = lstTradeItem.ListIndex
    Index = lstTradeItem.ListIndex + 1
    
    If Index < 1 Or Index > MAX_SHOPS Then Exit Sub
    
    With Shop(EditorIndex).TradeItem(Index)
        .item = cmbItem.ListIndex
        .ItemValue = Val(txtItemValue.text)
        .CostItem = cmbCostItem.ListIndex
        .CostItem2 = cmbCostItem2.ListIndex
        .CostValue = Val(txtCostValue.text)
        .CostValue2 = Val(txtCostValue2.text)
    End With
    UpdateShopTrade tmpPos
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdUpdate_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdDeleteTrade_Click()
    Dim Index As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Index = lstTradeItem.ListIndex + 1
    
    If Index < 1 Or Index > MAX_SHOPS Then Exit Sub
    
    With Shop(EditorIndex).TradeItem(Index)
        .item = 0
        .ItemValue = 0
        .CostItem = 0
        .CostItem2 = 0
        .CostValue = 0
        .CostValue2 = 0
    End With
    Call UpdateShopTrade
    lstTradeItem.ListIndex = Index
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdDeleteTrade_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Activate()
    hwndLastActiveWnd = hWnd
    If FormVisible("frmAdmin") And adminMin Then
        frmAdmin.centerMiniVert Width, Height, Left, Top
    End If
End Sub

Private Sub lstIndex_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ShopEditorInit
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstTradeItem_DblClick()
    Dim Index As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Index = lstTradeItem.ListIndex + 1
    
    If Index < 1 Or Index > MAX_SHOPS Then Exit Sub
    
    With Shop(EditorIndex).TradeItem(Index)
         cmbItem.ListIndex = .item
         txtItemValue.text = .ItemValue
         cmbCostItem.ListIndex = .CostItem
         cmbCostItem2.ListIndex = .CostItem2
         txtCostValue.text = .CostValue
         txtCostValue2.text = .CostValue2
    End With
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lstTradeItem_DblClick", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlBuy_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblBuy.Caption = "Buy Rate: " & scrlBuy.Value & "%"
    Shop(EditorIndex).BuyRate = scrlBuy.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlBuy_Change", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlSell_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblSell.Caption = "Sell Rate: " & scrlSell.Value & "%"
    Shop(EditorIndex).SellRate = scrlSell.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlSell_Change", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtCostValue_Change()
    Dim Index As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Index = lstTradeItem.ListIndex + 1
    
    If Index < 1 Or Index > MAX_SHOPS Then Exit Sub
    
    With Shop(EditorIndex).TradeItem(Index)
        If Not IsNumeric(txtCostValue.text) Then txtCostValue.text = 0
        If txtCostValue.text > MAX_LONG Then txtCostValue.text = MAX_LONG
        If txtCostValue.text < 0 Then txtCostValue.text = 0
    End With
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtCostValue_Change", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtCostValue2_Change()
    Dim Index As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Index = lstTradeItem.ListIndex + 1
    
    If Index < 1 Or Index > MAX_SHOPS Then Exit Sub
    
    With Shop(EditorIndex).TradeItem(Index)
        If Not IsNumeric(txtCostValue2.text) Then txtCostValue2.text = 0
        If txtCostValue2.text > MAX_LONG Then txtCostValue2.text = MAX_LONG
        If txtCostValue2.text < 0 Then txtCostValue2.text = 0
    End With
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtCostValue2_Change", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtItemValue_Change()
    Dim Index As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Index = lstTradeItem.ListIndex + 1
    
    If Index < 1 Or Index > MAX_SHOPS Then Exit Sub
    
    With Shop(EditorIndex).TradeItem(Index)
        If Not IsNumeric(txtItemValue.text) Then txtItemValue.text = 0
        If txtItemValue.text > MAX_LONG Then txtItemValue.text = MAX_LONG
        If txtItemValue.text < 0 Then txtItemValue.text = 0
    End With
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtItemValue_Change", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    TmpIndex = lstIndex.ListIndex
    Shop(EditorIndex).name = Trim$(txtName.text)
    
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Shop(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Max values
    txtName.MaxLength = NAME_LENGTH
    txtSearch.MaxLength = NAME_LENGTH
    
    ' Clear combo boxes
    frmEditor_Shop.cmbItem.Clear
    frmEditor_Shop.cmbItem.AddItem "None"
    frmEditor_Shop.cmbCostItem.Clear
    frmEditor_Shop.cmbCostItem.AddItem "None"
    frmEditor_Shop.cmbCostItem2.Clear
    frmEditor_Shop.cmbCostItem2.AddItem "None"

    For i = 1 To MAX_ITEMS
        frmEditor_Shop.cmbItem.AddItem i & ": " & Trim$(item(i).name)
        frmEditor_Shop.cmbCostItem.AddItem i & ": " & Trim$(item(i).name)
        frmEditor_Shop.cmbCostItem2.AddItem i & ": " & Trim$(item(i).name)
    Next
    
    frmEditor_Shop.cmbItem.ListIndex = 0
    frmEditor_Shop.cmbCostItem.ListIndex = 0
    frmEditor_Shop.cmbCostItem2.ListIndex = 0
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "Form_Load", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorSave = False Then
        ShopEditorCancel
    Else
        EditorSave = False
    End If
    frmAdmin.chkEditor(EDITOR_SHOP).Value = False
    BringWindowToTop (frmAdmin.hWnd)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "Form_Unload", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtName.SelStart = Len(txtName)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtName_GotFocus", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    HandleError "txtSearch_Change", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtSearch.SelStart = Len(txtSearch)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtSearch_GotFocus", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    HandleError "Form_KeyPress", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCopy_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    TmpIndex = lstIndex.ListIndex
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdCopy_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdPaste_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
      
    lstIndex.RemoveItem EditorIndex - 1
    Call CopyMemory(ByVal VarPtr(Shop(EditorIndex)), ByVal VarPtr(Shop(TmpIndex + 1)), LenB(Shop(TmpIndex + 1)))
    lstIndex.AddItem EditorIndex & ": " & Trim$(Shop(EditorIndex).name), EditorIndex - 1
    lstIndex.ListIndex = EditorIndex - 1
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdPaste_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
