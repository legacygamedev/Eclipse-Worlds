VERSION 5.00
Begin VB.Form frmEditor_Map 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14985
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Map.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   503
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   999
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picAttributes 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   7320
      ScaleHeight     =   7695
      ScaleWidth      =   7575
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton cmdCancel2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3960
         TabIndex        =   68
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Frame fraSoundEffect 
         Caption         =   "Sound Effect"
         Height          =   2655
         Left            =   2040
         TabIndex        =   101
         Top             =   2160
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdSoundEffect 
            Caption         =   "Accept"
            Height          =   375
            Left            =   240
            TabIndex        =   104
            Top             =   2040
            Width           =   1215
         End
         Begin VB.ComboBox cmbSoundEffect 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":038A
            Left            =   240
            List            =   "frmEditor_Map.frx":038C
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame fraSlide 
         Caption         =   "Slide"
         Height          =   2655
         Left            =   2040
         TabIndex        =   65
         Top             =   2160
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdSlide 
            Caption         =   "Accept"
            Height          =   375
            Left            =   240
            TabIndex        =   66
            Top             =   2040
            Width           =   1215
         End
         Begin VB.ComboBox cmbSlide 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":038E
            Left            =   240
            List            =   "frmEditor_Map.frx":039E
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame fraResource 
         Caption         =   "Resource"
         Height          =   2655
         Left            =   2040
         TabIndex        =   31
         Top             =   2160
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdResourceOk 
            Caption         =   "Accept"
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   2040
            Width           =   1215
         End
         Begin VB.HScrollBar scrlResource 
            Height          =   255
            Left            =   240
            Max             =   100
            Min             =   1
            TabIndex        =   33
            Top             =   480
            Value           =   1
            Width           =   2895
         End
         Begin VB.Label lblResource 
            Caption         =   "None"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame fraMapItem 
         Caption         =   "Item"
         Height          =   2655
         Left            =   2040
         TabIndex        =   40
         Top             =   2160
         Visible         =   0   'False
         Width           =   3375
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   2640
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   70
            Top             =   510
            Width           =   540
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00404040&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   15
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   71
               Top             =   15
               Width           =   480
               Begin VB.PictureBox picMapItem 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   480
                  Left            =   0
                  ScaleHeight     =   32
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   32
                  TabIndex        =   72
                  Top             =   0
                  Width           =   480
               End
            End
         End
         Begin VB.CommandButton cmdMapItem 
            Caption         =   "Accept"
            Height          =   375
            Left            =   240
            TabIndex        =   44
            Top             =   2040
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapItemValue 
            Height          =   255
            Left            =   240
            Min             =   1
            TabIndex        =   43
            Top             =   840
            Value           =   1
            Width           =   2295
         End
         Begin VB.HScrollBar scrlMapItem 
            Height          =   255
            Left            =   240
            Max             =   10
            Min             =   1
            TabIndex        =   42
            Top             =   480
            Value           =   1
            Width           =   2295
         End
         Begin VB.Label lblMapItem 
            Caption         =   "None"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame fraHeal 
         Caption         =   "Heal"
         Height          =   2655
         Left            =   2040
         TabIndex        =   56
         Top             =   2160
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbHeal 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":03B9
            Left            =   240
            List            =   "frmEditor_Map.frx":03C3
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton cmdHeal 
            Caption         =   "Accept"
            Height          =   375
            Left            =   240
            TabIndex        =   58
            Top             =   2040
            Width           =   1215
         End
         Begin VB.HScrollBar scrlHeal 
            Height          =   255
            Left            =   240
            Min             =   1
            TabIndex        =   57
            Top             =   960
            Value           =   1
            Width           =   2895
         End
         Begin VB.Label lblHeal 
            Caption         =   "Amount: 1"
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   720
            Width           =   2535
         End
      End
      Begin VB.Frame fraTrap 
         Caption         =   "Trap"
         Height          =   2655
         Left            =   2040
         TabIndex        =   61
         Top             =   2160
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbTrap 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":03D5
            Left            =   240
            List            =   "frmEditor_Map.frx":03DF
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   360
            Width           =   2895
         End
         Begin VB.HScrollBar scrlDamage 
            Height          =   255
            Left            =   240
            Min             =   1
            TabIndex        =   63
            Top             =   960
            Value           =   1
            Width           =   2895
         End
         Begin VB.CommandButton cmdTrap 
            Caption         =   "Accept"
            Height          =   375
            Left            =   240
            TabIndex        =   62
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblDamage 
            Caption         =   "Amount: 1"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   720
            Width           =   2535
         End
      End
      Begin VB.Frame fraShop 
         Caption         =   "Shop"
         Height          =   2655
         Left            =   2040
         TabIndex        =   53
         Top             =   2160
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdShop 
            Caption         =   "Accept"
            Height          =   375
            Left            =   240
            TabIndex        =   55
            Top             =   2040
            Width           =   1215
         End
         Begin VB.ComboBox cmbShop 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":03F1
            Left            =   240
            List            =   "frmEditor_Map.frx":03F3
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame fraNpcSpawn 
         Caption         =   "Npc Spawn"
         Height          =   2655
         Left            =   2040
         TabIndex        =   35
         Top             =   2160
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ListBox lstNpc 
            Height          =   780
            Left            =   240
            TabIndex        =   39
            Top             =   360
            Width           =   2895
         End
         Begin VB.HScrollBar scrlNpcDir 
            Height          =   255
            Left            =   240
            Max             =   3
            TabIndex        =   37
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton cmdNpcSpawn 
            Caption         =   "Accept"
            Height          =   375
            Left            =   240
            TabIndex        =   36
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblNpcDir 
            Caption         =   "Direction: Up"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   1320
            Width           =   2535
         End
      End
      Begin VB.Frame fraMapWarp 
         Caption         =   "Warp"
         Height          =   2655
         Left            =   2040
         TabIndex        =   45
         Top             =   2160
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdMapWarp 
            Caption         =   "Accept"
            Height          =   375
            Left            =   240
            TabIndex        =   52
            Top             =   2040
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapWarpY 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   51
            Top             =   1680
            Width           =   2895
         End
         Begin VB.HScrollBar scrlMapWarpX 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   49
            Top             =   1080
            Width           =   2895
         End
         Begin VB.HScrollBar scrlMapWarp 
            Height          =   255
            Left            =   240
            Min             =   1
            TabIndex        =   47
            Top             =   480
            Value           =   1
            Width           =   2895
         End
         Begin VB.Label lblMapWarpY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   1440
            Width           =   2895
         End
         Begin VB.Label lblMapWarpX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label lblMapWarp 
            Caption         =   "Map: 1"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   240
            Width           =   2895
         End
      End
   End
   Begin VB.CheckBox chkGrid 
      Caption         =   "Grid"
      Height          =   255
      Left            =   3360
      TabIndex        =   105
      ToolTipText     =   "Will place tiles you select randomly."
      Top             =   6000
      Width           =   675
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5280
      Left            =   120
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   352
      TabIndex        =   103
      Top             =   120
      Width           =   5280
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2280
      TabIndex        =   93
      Top             =   7080
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   7080
      Width           =   1020
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Properties"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   7080
      Width           =   1020
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   7080
      Width           =   1020
   End
   Begin VB.VScrollBar scrlPictureY 
      Height          =   5295
      Left            =   5400
      Max             =   255
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar scrlPictureX 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   5295
   End
   Begin VB.Frame fraType 
      Caption         =   "Type"
      Height          =   1575
      Left            =   4320
      TabIndex        =   86
      Top             =   5880
      Width           =   1455
      Begin VB.OptionButton OptLayers 
         Alignment       =   1  'Right Justify
         Caption         =   "Layers"
         Height          =   255
         Left            =   360
         TabIndex        =   99
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptAttributes 
         Alignment       =   1  'Right Justify
         Caption         =   "Attributes"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton OptBlock 
         Alignment       =   1  'Right Justify
         Caption         =   "Block"
         Height          =   255
         Left            =   480
         TabIndex        =   97
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton OptEvents 
         Alignment       =   1  'Right Justify
         Caption         =   "Events"
         Height          =   255
         Left            =   360
         TabIndex        =   96
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chkShowAttributes 
         Alignment       =   1  'Right Justify
         Caption         =   "Attributes"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         ToolTipText     =   "Will show the attribute's text on the map."
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1115
      End
   End
   Begin VB.Frame fraRandom 
      Caption         =   "Random"
      Height          =   1575
      Left            =   4320
      TabIndex        =   76
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   800
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   88
         Top             =   240
         Width           =   540
         Begin VB.PictureBox Picture8 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   89
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picRandomTile 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   1
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   90
               Top             =   0
               Width           =   480
            End
         End
      End
      Begin VB.PictureBox Picture11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   800
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   83
         Top             =   900
         Width           =   540
         Begin VB.PictureBox Picture12 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   84
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picRandomTile 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   3
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   85
               Top             =   0
               Width           =   480
            End
         End
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   120
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   80
         Top             =   900
         Width           =   540
         Begin VB.PictureBox Picture10 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   81
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picRandomTile 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   2
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   82
               Top             =   0
               Width           =   480
            End
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   120
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   77
         Top             =   240
         Width           =   540
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   78
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picRandomTile 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   0
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   79
               Top             =   0
               Width           =   480
            End
         End
      End
   End
   Begin VB.Frame fraTileSet 
      Caption         =   "Tileset: 0"
      Height          =   1215
      Left            =   120
      TabIndex        =   27
      Top             =   5760
      Width           =   4095
      Begin VB.CheckBox chkTilePreview 
         Caption         =   "Tile Preview"
         Height          =   255
         Left            =   240
         TabIndex        =   91
         Top             =   240
         Width           =   1275
      End
      Begin VB.CheckBox chkEyeDropper 
         Caption         =   "Eye Dropper"
         Height          =   255
         Left            =   1680
         TabIndex        =   75
         ToolTipText     =   "Will find the tile on the layer you select."
         Top             =   240
         Width           =   1275
      End
      Begin VB.CheckBox chkRandom 
         Caption         =   "Random"
         Height          =   255
         Left            =   1680
         TabIndex        =   73
         ToolTipText     =   "Will place tiles you select randomly."
         Top             =   480
         Width           =   915
      End
      Begin VB.HScrollBar scrlTileSet 
         Height          =   255
         Left            =   240
         Max             =   10
         Min             =   1
         TabIndex        =   2
         Top             =   840
         Value           =   1
         Width           =   3615
      End
      Begin VB.Label lblRevision 
         BackStyle       =   0  'Transparent
         Caption         =   "Revision:"
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   600
         Width           =   3615
      End
   End
   Begin VB.Frame fraLayers 
      Caption         =   "Layers"
      Height          =   5775
      Left            =   5760
      TabIndex        =   29
      Top             =   0
      Width           =   1455
      Begin VB.HScrollBar scrlAutotile 
         Height          =   255
         Left            =   240
         Max             =   5
         TabIndex        =   94
         Top             =   4440
         Width           =   975
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Fringe"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Ground"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Roof"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Cover"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   5280
         Width           =   975
      End
      Begin VB.CommandButton cmdFill 
         Caption         =   "Fill"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label lblAutoTile 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   4200
         Width           =   1215
      End
   End
   Begin VB.Frame fraAttribs 
      Caption         =   "Attributes"
      Height          =   5775
      Left            =   5760
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      Begin VB.OptionButton optSound 
         Caption         =   "Sound"
         Height          =   270
         Left            =   120
         TabIndex        =   100
         Top             =   3360
         Width           =   1215
      End
      Begin VB.OptionButton optGravity 
         Caption         =   "Gravity"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   3120
         Width           =   1225
      End
      Begin VB.CommandButton cmdAttributeFill 
         Caption         =   "Fill"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   4800
         Width           =   975
      End
      Begin VB.OptionButton optCheckpoint 
         Caption         =   "Checkpoint"
         Height          =   270
         Left            =   120
         TabIndex        =   17
         Top             =   2880
         Width           =   1215
      End
      Begin VB.OptionButton optSlide 
         Caption         =   "Slide"
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   1215
      End
      Begin VB.OptionButton optTrap 
         Caption         =   "Trap"
         Height          =   270
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton optHeal 
         Caption         =   "Heal"
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton optBank 
         Caption         =   "Bank"
         Height          =   270
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton optShop 
         Caption         =   "Shop"
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optNpcSpawn 
         Caption         =   "Npc Spawn"
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optResource 
         Caption         =   "Resource"
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optBlocked 
         Caption         =   "Blocked"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optWarp 
         Caption         =   "Warp"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdAttributeClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   5280
         Width           =   975
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optNpcAvoid 
         Caption         =   "Npc Avoid"
         Height          =   270
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmEditor_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkEyeDropper_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlAutotile.Value = 0
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkEyeDropper_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkTilePreview_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    CurX = 0
    CurY = 0
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkTilePreview_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbHeal_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorVitalType = cmbHeal.ListIndex + 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbHeal_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
   ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If AlertMsg("Are you sure you want to erase this map?", False, False) = YES Then
        Call ClearMap
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdSoundEffect_Click()
   ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If cmbSoundEffect.ListIndex = 0 Then Exit Sub
    
    MapEditorSound = SoundCache(cmbSoundEffect.ListIndex)
    picAttributes.Visible = False
    fraSoundEffect.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSoundEffect_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optSound_Click()
   ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeFrames
    picAttributes.Visible = True
    fraSoundEffect.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optSound_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub picBack_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    x = x + (frmEditor_Map.scrlPictureX.Value * PIC_X)
    y = y + (frmEditor_Map.scrlPictureY.Value * PIC_Y)
    
    Call MapEditorChooseTile(Button, x, y)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picBack_MouseDown", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    x = x + (frmEditor_Map.scrlPictureX.Value * PIC_X)
    y = y + (frmEditor_Map.scrlPictureY.Value * PIC_Y)
    
    If scrlAutotile.Value = 0 Then
        Call MapEditorDrag(Button, x, y)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picBack_MouseMove", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlAutotile_Change()
    Select Case scrlAutotile.Value
        Case 0 ' Normal
            lblAutoTile.Caption = "Normal"
        Case 1 ' Autotile
            lblAutoTile.Caption = "Autotile"
        Case 2 ' Fake autotile
            lblAutoTile.Caption = "Fake"
        Case 3 ' Animated
            lblAutoTile.Caption = "Animated"
        Case 4 ' Cliff
            lblAutoTile.Caption = "Cliff"
        Case 5 ' Waterfall
            lblAutoTile.Caption = "Waterfall"
    End Select
    
    SetMapAutotileScrollbar
End Sub

Private Sub cmbShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
        
    EditorShop = cmbShop.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbShop_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbTrap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorVitalType = cmbTrap.ListIndex + 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbTrap_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCancel2_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeFrames
    picAttributes.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdCancel2_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdHeal_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorVitalType = cmbHeal.ListIndex + 1
    MapEditorVitalAmount = scrlHeal.Value
    picAttributes.Visible = False
    fraHeal.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdHeal_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdMapItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ItemEditorNum = scrlMapItem.Value
    ItemEditorValue = scrlMapItemValue.Value
    picAttributes.Visible = False
    fraMapItem.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdMapItem_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdMapWarp_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    EditorWarpMap = scrlMapWarp.Value
    EditorWarpX = scrlMapWarpX.Value
    EditorWarpY = scrlMapWarpY.Value
    picAttributes.Visible = False
    fraMapWarp.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdMapWarp_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdNpcSpawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    SpawnNpcNum = lstNpc.ListIndex + 1
    SpawnNpcDir = scrlNpcDir.Value
    picAttributes.Visible = False
    fraNpcSpawn.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdNpcSpawn_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdResourceOk_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ResourceEditorNum = scrlResource.Value
    picAttributes.Visible = False
    fraResource.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdResourceOk_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorShop = cmbShop.ListIndex
    picAttributes.Visible = False
    fraShop.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdShop_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdSlide_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorSlideDir = cmbSlide.ListIndex
    picAttributes.Visible = False
    fraSlide.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSlide_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdTrap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorVitalType = cmbTrap.ListIndex + 1
    MapEditorVitalAmount = scrlDamage.Value
    picAttributes.Visible = False
    fraTrap.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdTrap_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAttributeFill_Click()
    Dim Button As Integer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call MapEditorFillAttributes(Button)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdAttributeFill_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Move the entire attributes box on screen
    picAttributes.Left = 0
    picAttributes.Top = 0
    
    ' Set maxes for attribute forms
    scrlMapItem.max = MAX_ITEMS
    scrlResource.max = MAX_RESOURCES
    scrlMapWarp.max = MAX_MAPS
    
    ' Set the width of the form
    Me.Width = 7380
    
    ' Set the max scrollbar to the number of tilesets
    frmEditor_Map.scrlTileSet.max = NumTileSets
    
    ' Populate the cache if we need to
    If Not HasPopulated Then
        PopulateLists
    End If
    
    ' Add the array to the combo
    frmEditor_Map.cmbSoundEffect.Clear
    frmEditor_Map.cmbSoundEffect.AddItem "None"

    For i = 1 To UBound(SoundCache)
        frmEditor_Map.cmbSoundEffect.AddItem SoundCache(i)
    Next
    
    frmEditor_Map.cmbSoundEffect.ListIndex = 0
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Load", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorSave = False Then
        MapEditorCancel
    Else
        EditorSave = False
    End If
    Call ToggleGUI(True)
    
    ' Make sure the properties form is closed
    If frmEditor_MapProperties.Visible Then
        Unload frmEditor_MapProperties
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Unload", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optBlock_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    picBack.Visible = True
    scrlPictureY.Visible = True
    scrlPictureX.Visible = True
    frmEditor_Map.chkEyeDropper.Enabled = True
    frmEditor_Map.chkRandom.Enabled = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optBlock_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optHeal_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    cmbHeal.ListIndex = 0
    ClearAttributeFrames
    picAttributes.Visible = True
    fraHeal.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optHeal_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optLayer_Click(Index As Integer)
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Set which layer we're on
    CurrentLayer = Index
    
    If chkRandom = 1 Then
        EditorTileX = 1
        EditorTileY = 1
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optLayer_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optLayers_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If OptLayers.Value Then
        fraLayers.Visible = True
        fraAttribs.Visible = False
    End If
    
    chkEyeDropper.Enabled = True
    frmEditor_Map.chkRandom.Enabled = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optLayers_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optAttributes_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If OptAttributes.Value Then
        fraLayers.Visible = False
        fraAttribs.Visible = True
    End If
    
    frmEditor_Map.chkEyeDropper.Enabled = True
    frmEditor_Map.chkRandom.Enabled = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optAttribs_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optNpcSpawn_Click()
    Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lstNpc.Clear
    
    For n = 1 To MAX_MAP_NPCS
        If Map.NPC(n) > 0 Then
            lstNpc.AddItem n & ": " & NPC(Map.NPC(n)).name
        Else
            lstNpc.AddItem n & ": No Npc"
        End If
    Next n
    
    scrlNpcDir.Value = 0
    lstNpc.ListIndex = 0
    
    ClearAttributeFrames
    picAttributes.Visible = True
    fraNpcSpawn.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optNpcSpawn_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkRandom_Click()
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmEditor_Map.fraRandom.Visible = Not frmEditor_Map.fraRandom.Visible
    frmEditor_Map.fraType.Visible = Not frmEditor_Map.fraType.Visible
    fraLayers.Visible = True
    fraAttribs.Visible = False
    frmEditor_Map.OptLayers.Value = True
    
    If frmEditor_Map.chkRandom = 1 Then
        EditorTileX = 1
        EditorTileY = 1
        frmEditor_Map.optLayer(MapLayer.Ground).Value = 1
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkRandom_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optResource_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeFrames
    If Not Trim$(Resource(scrlResource.Value).name) = vbNullString Then
        lblResource.Caption = Trim$(Resource(scrlResource.Value).name)
    End If
    picAttributes.Visible = True
    fraResource.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optResource_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeFrames
    picAttributes.Visible = True
    fraShop.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optShop_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optSlide_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    cmbSlide.ListIndex = 0
    ClearAttributeFrames
    picAttributes.Visible = True
    fraSlide.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optSlide_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optSprite_Click()
  ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
   
    ClearAttributeFrames
    picAttributes.Visible = True
    Exit Sub
   
' Error handler
ErrorHandler:
    HandleError "optSprite_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optTrap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    cmbTrap.ListIndex = 0
    ClearAttributeFrames
    picAttributes.Visible = True
    fraTrap.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optTrap_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapEditorDrag(Button As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Button = vbLeftButton Then
        ' Convert the pixel number to tile number
        x = (x \ PIC_X) + 1
        y = (y \ PIC_Y) + 1
        
        ' Check it's not out of bounds
        If x < 0 Then x = 0
        If x > Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width / PIC_X Then x = Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width / PIC_X
        If y < 0 Then y = 0
        If y > Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height / PIC_Y Then y = Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height / PIC_Y
        
        ' Find out what to set the width + height of map editor to
        If x > EditorTileX Then ' Drag right
            EditorTileWidth = x - EditorTileX
        Else ' Drag left
            ' TO DO
        End If
        If y > EditorTileY Then ' Drag down
            EditorTileHeight = y - EditorTileY
        Else ' Drag up
            ' TO DO
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapEditorDrag", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorSave = True
    Call MapEditorSave
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCancel_Click()
    Dim Result As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Unload frmEditor_Map
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdCancel_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdProperties_Click()
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Load the values
    MapPropertiesInit
    
    ' Update the 1stnpcs list Index so it is selected
    frmEditor_MapProperties.lstNpcs.ListIndex = 0
    
    ' Show the form
    frmEditor_MapProperties.Show
    
    ' Lock map editor open til map properties is closed
    frmEditor_Map.cmdSave.Enabled = False
    frmEditor_Map.cmdCancel.Enabled = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdProperties_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optWarp_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeFrames
    picAttributes.Visible = True
    fraMapWarp.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optWarp_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeFrames
    picAttributes.Visible = True
    fraMapItem.Visible = True
    
    If Not Trim$(Item(scrlMapItem.Value).name) = vbNullString Then
        lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).name) & " x" & scrlMapItemValue.Value
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optItem_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdFill_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorFillLayer
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdFill_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdClear_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call MapEditorClearLayer
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdClear_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAttributeClear_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call MapEditorClearAttributes
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdAttributeClear_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picRandomTile_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    RandomTileSelected = Index
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picRandomTile_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlHeal_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorVitalAmount = scrlHeal.Value
    lblHeal.Caption = "Amount: " & scrlHeal.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlHeal_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorVitalAmount = scrlDamage.Value
    lblDamage.Caption = "Amount: " & scrlDamage.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlDamage_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapItem_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Item(scrlMapItem.Value).Type = ITEM_TYPE_CURRENCY Then
        scrlMapItemValue.Enabled = True
    Else
        scrlMapItemValue.Value = 1
        scrlMapItemValue.Enabled = False
    End If
    
    If Not Trim$(Item(scrlMapItem.Value).name) = vbNullString Then
        lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).name) & " x" & scrlMapItemValue.Value
    Else
        lblMapItem.Caption = "None"
        frmEditor_Map.picMapItem.Cls
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapItem_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapItem_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlMapItem_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapItem_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapItemValue_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).name) & " x" & scrlMapItemValue.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapItemValue_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapItemValue_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlMapItemValue_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapItemValue_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapWarp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMapWarp.Caption = "Map: " & scrlMapWarp.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapWarp_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapWarp_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlMapWarp_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapWarp_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapWarpX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMapWarpX.Caption = "X: " & scrlMapWarpX.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapWarpX_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapWarpX_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlMapWarpX_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapWarpX_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapWarpY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMapWarpY.Caption = "Y: " & scrlMapWarpY.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapWarpY_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapWarpY_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlMapWarpY_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapWarpY_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlNpcDir_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case scrlNpcDir.Value
        Case DIR_UP
            lblNpcDir = "Direction: Up"
        Case DIR_DOWN
            lblNpcDir = "Direction: Down"
        Case DIR_LEFT
            lblNpcDir = "Direction: Left"
        Case DIR_RIGHT
            lblNpcDir = "Direction: Right"
    End Select
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlNpcDir_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlNpcDir_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlNpcDir_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlNpcDir_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlResource_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not Trim$(Resource(scrlResource.Value).name) = vbNullString Then
        lblResource.Caption = Resource(scrlResource.Value).name
    Else
        lblResource.Caption = "None"
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlResource_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlResource_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlResource_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlResource_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlPictureX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call MapEditorTileScroll
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlPictureX_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlPictureY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call MapEditorTileScroll
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlPictureY_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlPictureX_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlPictureX_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlPictureX_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlPictureY_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlPictureY_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlPictureY_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlTileSet_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    fraTileSet.Caption = "Tileset: " & scrlTileSet.Value
    
    frmEditor_Map.scrlPictureY.max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height \ PIC_Y) - (frmEditor_Map.picBack.Height \ PIC_Y)
    frmEditor_Map.scrlPictureX.max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width \ PIC_X) - (frmEditor_Map.picBack.Width \ PIC_X)
    
    MapEditorTileScroll
    
    EditorTileX = 0
    EditorTileY = 0
    EditorTileWidth = 1
    EditorTileHeight = 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlTileSet_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlTileSet_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlTileSet_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlTileSet_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
