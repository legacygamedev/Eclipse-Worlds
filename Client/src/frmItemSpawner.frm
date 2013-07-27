VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemSpawner 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Item Spawner - Equipment -> 47 items"
   ClientHeight    =   4935
   ClientLeft      =   8280
   ClientTop       =   4425
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   553
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList itemsImageList 
      Left            =   7710
      Top             =   4290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CheckBox chkClose 
      Caption         =   "Close window after spawning"
      ForeColor       =   &H00C0C000&
      Height          =   420
      Left            =   4410
      TabIndex        =   10
      Top             =   285
      Value           =   1  'Checked
      Width           =   1425
   End
   Begin MSComctlLib.ListView listItems 
      Height          =   3795
      Left            =   105
      TabIndex        =   9
      Top             =   1170
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   6694
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      HotTracking     =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483634
      BorderStyle     =   1
      Appearance      =   0
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdSpawn 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      Caption         =   "Spawn it"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3270
      TabIndex        =   8
      Top             =   345
      Width           =   795
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2505
      TabIndex        =   7
      Text            =   "1"
      Top             =   285
      Width           =   705
   End
   Begin VB.OptionButton radioInv 
      Caption         =   "Inventory(3 slots)"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   525
      Value           =   -1  'True
      Width           =   1515
   End
   Begin VB.OptionButton radioGround 
      Caption         =   "Ground"
      Enabled         =   0   'False
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   855
   End
   Begin MSComctlLib.TabStrip tabItems 
      Height          =   4185
      Left            =   15
      TabIndex        =   0
      Top             =   780
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   7382
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   10
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Recent"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "None"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Equipment"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Consumable"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Title"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Spell"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Teleport"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Reset Stats"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Auto Life"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sprite Changer"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2625
      ScaleHeight     =   270
      ScaleWidth      =   2685
      TabIndex        =   12
      Top             =   2610
      Width           =   2685
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Caption         =   "No items available in this category!"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   15
         TabIndex        =   13
         Top             =   30
         Width           =   2655
      End
   End
   Begin VB.Label lblHelp2 
      BackStyle       =   0  'Transparent
      Caption         =   "and ""Spawn It""."
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   6075
      TabIndex        =   14
      Top             =   480
      Width           =   1020
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblOptions 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   195
      Left            =   4635
      TabIndex        =   11
      Top             =   15
      Width           =   645
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   3
      X1              =   305
      X2              =   382.333
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line lineAmount 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      X1              =   174
      X2              =   251.333
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Label lblAmount 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   195
      Left            =   2910
      TabIndex        =   6
      Top             =   30
      Width           =   645
   End
   Begin VB.Label lblHelp1 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose the item, input Amount"
      BeginProperty Font 
         Name            =   "Myriad Arabic"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   6000
      TabIndex        =   5
      Top             =   270
      Width           =   2325
      WordWrap        =   -1  'True
   End
   Begin VB.Line lineHow 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      X1              =   402
      X2              =   546
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Label lblHow 
      Alignment       =   1  'Right Justify
      Caption         =   "How to use it"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   7020
      TabIndex        =   4
      Top             =   0
      Width           =   1185
   End
   Begin VB.Line lineWhere 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   144
      Y1              =   14
      Y2              =   14
   End
   Begin VB.Label lblWhere 
      Caption         =   "Where to Spawn It"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1665
   End
End
Attribute VB_Name = "frmItemSpawner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lastTab As Byte
Private allowTitle As Boolean
Private currentItemId As Long
Private currentAmount As Long
Private picked As Boolean
Private Declare Function SendMessage Lib "user32" Alias _
 "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
 ByVal wParam As Long, lParam As Any) As Long
 
Public Function ListView_SetIconSpacing(hWndLV As Long, cx As Long, cy As Long) As Long
    Dim LVM_SETICONSPACING As Long
    LVM_SETICONSPACING = 4149
  ListView_SetIconSpacing = SendMessage(hWndLV, LVM_SETICONSPACING, 0, ByVal MakeLong(cx, cy))
End Function

Private Sub styleListwView(sType As status, Optional Msg As String)
    
    Select Case sType
    
        Case status.Correct
            listItems.BackColor = &H8000000E
            listItems.BorderStyle = ccFixedSingle
            picInfo.Visible = False
            picInfo.ZOrder 1
        Case status.Error
            listItems.BackColor = &H8000000F
            listItems.BorderStyle = ccNone
            picInfo.Visible = True
            picInfo.ZOrder 0
            lblInfo.Caption = Msg
    End Select

End Sub

Private Function generateItemsForTab(tabNum As Byte) As Boolean
Dim i As Long, z As Long, tempItems() As ItemRec, ret As Boolean
    tabNum = tabNum - 2
    
    Select Case tabNum
        Case ITEM_TYPE_NONE
            ret = populateSpecificType(tempItems, ITEM_TYPE_NONE)
        Case ITEM_TYPE_EQUIPMENT
            ret = populateSpecificType(tempItems, ITEM_TYPE_EQUIPMENT)
        Case ITEM_TYPE_CONSUME
            ret = populateSpecificType(tempItems, ITEM_TYPE_CONSUME)
        Case ITEM_TYPE_TITLE
            ret = populateSpecificType(tempItems, ITEM_TYPE_TITLE)
        Case ITEM_TYPE_SPELL
            ret = populateSpecificType(tempItems, ITEM_TYPE_SPELL)
       Case ITEM_TYPE_TELEPORT
            ret = populateSpecificType(tempItems, ITEM_TYPE_TELEPORT)
       Case ITEM_TYPE_RESETSTATS
            ret = populateSpecificType(tempItems, ITEM_TYPE_RESETSTATS)
       Case ITEM_TYPE_AUTOLIFE
            ret = populateSpecificType(tempItems, ITEM_TYPE_AUTOLIFE)
       Case ITEM_TYPE_SPRITE
            ret = populateSpecificType(tempItems, ITEM_TYPE_SPRITE)
    End Select
    If ret Then
        Set listItems.Icons = itemsImageList
                
        For i = 0 To UBound(tempItems)
            listItems.listItems.Add , , Trim(tempItems(i).name), itemsImageList.ListImages(i + 1).Index
        Next
        generateItemsForTab = True
    End If


End Function
Private Sub generateLastItems()
    If lastSpawnedItemsCounter = 0 Then
        styleListwView status.Error, "You haven't spawned any items yet!"
    Else
        styleListwView status.Correct
    End If
End Sub

Private Sub cmdSpawn_Click()

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If
    
    SendSpawnItem currentlyListedIndexes(currentItemId), CLng(txtAmount)
    
    If chkClose.Value = 1 Then
        Unload Me
        frmAdmin.lastIndex = -1
        lastTab = 0
        currentItemId = 0
    End If
    
    Exit Sub
' Error handler
errorhandler:
    HandleError "cmdSpawn_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Load()
    tabItems_Click
    ListView_SetIconSpacing listItems.hWnd, 105, 56
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAdmin.styleButtons
End Sub

Private Sub listItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdSpawn.Enabled = True
    currentItemId = Item.Index - 1
    picked = True
    Me.Caption = "Item Spawner - Going to spawn " & txtAmount.text & " " & listItems.listItems(Item.Index).text
End Sub

Private Sub tabItems_Click()
    If lastTab = tabItems.SelectedItem.Index Then Exit Sub
    cmdSpawn.Enabled = False
    listItems.listItems.Clear
    Set listItems.Icons = Nothing
    Set listItems.SmallIcons = Nothing
    itemsImageList.ListImages.Clear
        
    If tabItems.SelectedItem.Index = 1 Then
        generateLastItems
        Exit Sub
    End If
    
    If generateItemsForTab(tabItems.SelectedItem.Index) Then
        styleListwView status.Correct
    Else
        styleListwView status.Error, "No items available in this category!"
    End If
    frmAdmin.lastIndex = lastTab - 1
    frmAdmin.optCat(tabItems.SelectedItem.Index - 1).Value = True
    frmAdmin.optCat_MouseUp tabItems.SelectedItem.Index - 1, 0, 0, 0, 0
    
    Me.Caption = "Item Spawner - " & tabItems.SelectedItem.Caption & " -> " & listItems.listItems.count & " items available"
    
    lastTab = tabItems.SelectedItem.Index
End Sub
Private Function correctValue(ByRef textBox As textBox, ByRef valueToChange, min As Long, max As Long, Optional defaultVal As Long = 0) As Boolean
    Dim test As textBox, TempValue As String
    
    If textBox.text = "" Then
        textBox.text = CStr(defaultVal)
        valueToChange = defaultVal
        correctValue = True
    End If

    If Len(textBox.text) = 1 And InStr(1, textBox.text, "-") = 1 Then
        correctValue = True
        Exit Function
    ElseIf Len(textBox.text) = 1 And IsNumeric(textBox.text) Then
        If verifyValue(textBox, min, max) Then
            TempValue = textBox.text
            valueToChange = TempValue
            correctValue = True
        Else
            textBox.text = CStr(valueToChange)
            textBox.SelStart = Len(textBox.text)
            correctValue = False
        End If
    ElseIf Len(textBox.text) > 1 And InStr(1, textBox.text, "-") = 0 And InStrRev(textBox.text, "-") = 0 And IsNumeric(textBox.text) Then

        If verifyValue(textBox, min, max) Then
            TempValue = textBox.text
            valueToChange = TempValue
            correctValue = True
        Else
            textBox.text = CStr(valueToChange)
            textBox.SelStart = Len(textBox.text)
            correctValue = False
        End If

    ElseIf Len(textBox.text) > 1 And InStr(1, textBox.text, "-") = 1 And InStrRev(textBox.text, "-") = 1 And IsNumeric(textBox.text) Then

        If verifyValue(textBox, min, max) Then
            TempValue = textBox.text
            valueToChange = TempValue
            correctValue = True
        Else
            textBox.text = CStr(valueToChange)
            textBox.SelStart = Len(textBox.text)
        correctValue = False
        End If
        
    Else
        textBox.text = CStr(valueToChange)
        textBox.SelStart = Len(textBox.text)
        correctValue = False
    End If
End Function

Private Sub reviseValue(ByRef textBox As textBox, ByRef valueToChange)
    If Not IsNumeric(textBox.text) Then
        textBox.text = CStr(valueToChange)
    Else
        textBox.text = CStr(valueToChange)
    End If
End Sub

Private Function verifyValue(txtBox As textBox, min As Long, max As Long)
    Dim Msg As String
    
    If (CLng(txtBox.text) >= min And CLng(txtBox.text) <= max) Then
        verifyValue = True
    Else
        verifyValue = False
    End If
End Function
Private Sub selectValue(ByRef textBox As textBox)
    textBox.SelStart = 0
    textBox.SelLength = Len(textBox.text)
End Sub

Private Sub txtAmount_Change()
    If correctValue(txtAmount, currentAmount, 0, 20) Then
        If picked Then
                Me.Caption = "Item Spawner - Going to spawn " & txtAmount.text & " " & listItems.listItems(listItems.SelectedItem.Index).text
        End If
    End If
End Sub

Private Sub txtAmount_Click()
     selectValue txtAmount
End Sub

Private Sub txtAmount_GotFocus()
     selectValue txtAmount
End Sub

Private Sub txtAmount_LostFocus()
    reviseValue txtAmount, currentAmount
End Sub
