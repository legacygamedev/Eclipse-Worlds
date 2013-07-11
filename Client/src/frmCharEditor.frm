VERSION 5.00
Begin VB.Form frmCharEditor 
   Caption         =   "Character Editor"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameCharPanel 
      Caption         =   "Editing - None"
      Height          =   5415
      Left            =   2370
      TabIndex        =   0
      Top             =   210
      Width           =   7275
      Begin VB.TextBox txtLevel 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   630
         TabIndex        =   2
         Text            =   "0"
         Top             =   330
         Width           =   975
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   360
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmCharEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub frameCharPanel_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
