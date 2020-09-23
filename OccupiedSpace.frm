VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form OccupiedSpace 
   AutoRedraw      =   -1  'True
   Caption         =   "      Occupied Space concept 2         "
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8055
   Icon            =   "OccupiedSpace.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TemporizeToolTipRTF 
      Enabled         =   0   'False
      Left            =   360
      Top             =   1800
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1095
      Left            =   2760
      TabIndex        =   111
      TabStop         =   0   'False
      ToolTipText     =   "double click to close"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"OccupiedSpace.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton SuckInFocus 
      Caption         =   "SuckInFocus"
      Height          =   240
      Left            =   120
      TabIndex        =   105
      Top             =   1200
      Width           =   1830
   End
   Begin VB.Timer BlinkLED 
      Left            =   5520
      Top             =   2040
   End
   Begin VB.Frame CommandsFrame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   6975
      Begin VB.CommandButton DesignTimeButton 
         Caption         =   "design time "
         Height          =   495
         Left            =   0
         TabIndex        =   110
         ToolTipText     =   "run/design time objects dimensions"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton ToggleContainer 
         Caption         =   "container is picture"
         Height          =   495
         Left            =   1200
         TabIndex        =   109
         ToolTipText     =   "between Form and Picture"
         Top             =   120
         Width           =   855
      End
      Begin VB.CheckBox ViewCellsCheck 
         Caption         =   "view cells"
         Height          =   495
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton ClearButton 
         Caption         =   "clear"
         Height          =   495
         Left            =   2400
         TabIndex        =   4
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton AddButton 
         Caption         =   "add xxx objects        [of yyy remaining] "
         Height          =   495
         Left            =   3600
         TabIndex        =   3
         ToolTipText     =   "unsuccess in adding, lights LED temporarily"
         Top             =   120
         Width           =   1695
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   495
         Left            =   5640
         TabIndex        =   2
         Top             =   120
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   873
         _Version        =   393216
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000C0&
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   5340
         Shape           =   3  'Circle
         Top             =   255
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1440
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   108
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   99
      Left            =   0
      TabIndex        =   103
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   98
      Left            =   0
      TabIndex        =   102
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Occupied Space 2:             - variable size objects      - choice of container    - view elementary cells"
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   5160
      TabIndex        =   106
      ToolTipText     =   "This is just a comment label, visible only at design time"
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "<-- this is a pile of 100 labels (index=0 to 99)   YOU CAN MANUALLY CHANGE THEIR DIMENSIONS AND THEN RUN 'DESIGN TIME' MODE"
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   1080
      TabIndex        =   104
      ToolTipText     =   "This is just a comment label, visible only at design time"
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   97
      Left            =   0
      TabIndex        =   101
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   96
      Left            =   0
      TabIndex        =   100
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   95
      Left            =   0
      TabIndex        =   99
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   94
      Left            =   0
      TabIndex        =   98
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   93
      Left            =   0
      TabIndex        =   97
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   92
      Left            =   0
      TabIndex        =   96
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   91
      Left            =   0
      TabIndex        =   95
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   90
      Left            =   0
      TabIndex        =   94
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   89
      Left            =   0
      TabIndex        =   93
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   88
      Left            =   0
      TabIndex        =   92
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   87
      Left            =   0
      TabIndex        =   91
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   86
      Left            =   0
      TabIndex        =   90
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   85
      Left            =   0
      TabIndex        =   89
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   84
      Left            =   0
      TabIndex        =   88
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   83
      Left            =   0
      TabIndex        =   87
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   82
      Left            =   0
      TabIndex        =   86
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   81
      Left            =   0
      TabIndex        =   85
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   80
      Left            =   0
      TabIndex        =   84
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   79
      Left            =   0
      TabIndex        =   83
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   78
      Left            =   0
      TabIndex        =   82
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   77
      Left            =   0
      TabIndex        =   81
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   76
      Left            =   0
      TabIndex        =   80
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   75
      Left            =   0
      TabIndex        =   79
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   74
      Left            =   0
      TabIndex        =   78
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   73
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   72
      Left            =   0
      TabIndex        =   76
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   71
      Left            =   0
      TabIndex        =   75
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   70
      Left            =   0
      TabIndex        =   74
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   69
      Left            =   0
      TabIndex        =   73
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   68
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   67
      Left            =   0
      TabIndex        =   71
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   66
      Left            =   0
      TabIndex        =   70
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   65
      Left            =   0
      TabIndex        =   69
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   64
      Left            =   0
      TabIndex        =   68
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   63
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   62
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   61
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   60
      Left            =   0
      TabIndex        =   64
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   59
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   58
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   57
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   56
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   55
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   54
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   53
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   52
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   51
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   50
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   49
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   48
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   47
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   46
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   45
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   44
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   43
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   42
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   41
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   40
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   39
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   38
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   37
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   36
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   35
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   34
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   33
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   32
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   31
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   30
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   29
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   28
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Menu menu 
      Caption         =   ".   ?   ."
      Begin VB.Menu menuusage 
         Caption         =   "Usage"
      End
      Begin VB.Menu menuabout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu menucomment 
      Caption         =   " xxx objs random allocation, no collision"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "OccupiedSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Title of present application:
'- 'Occupied space 2' concept used to allocate 100 labels randomly w/o overlapping (no API).
'  submitted to www.planet-source-code.com/vb the 21st August 2K1, by Pietro Cecchi
'  [variable size labels (may be all different): elementary cell equals minimum width and minimum height of all 100 labels]

'Ideated by A. Gunnar, Realized by Pietro Cecchi
'Emails: g68_se@yahoo.se  pietrocecchi@inwind.it
'Posted on www.planet.source-code.com/vb August 25st 2K1
'Written down in Visual Basic 6

'Same argument applications posted by me before on the planet:
'- How to allocate HUNDREDS rectangles (labels, images...) without overlapping, no API!
'  submitted to www.planet-source-code.com/vb the 18th August 2K1, by Pietro Cecchi
'- Intersection of HUNDREDS rectangles, no API!
'  submitted to www.planet-source-code.com/vb the 15th August 2K1, by Pietro Cecchi
'- 'Occupied space' concept used to allocate 100 labels randomly w/o overlapping (no API).
'  submitted to www.planet-source-code.com/vb the 21st August 2K1, by Pietro Cecchi
'  [fixed size labels (all identical): a cell equals each label size]

'The whole application comes in one form and one module, no API required
''unoccupied space' concept is used, instead of intersection


Dim DesignTimeLabelSize As Boolean 'if true, labels will maintain their original (design time) size
                                   'if false, label size will be changed randomly (sizes vary in the integer range of 1 to 4)

Dim ChoosenContainer As Object    'can be a form or a picture
                                  'assigned in Form_Initialize

Dim MaxAllocObjsNumber As Integer 'max allocable objs, computed by OccupiedSpaceFun and
                                  'refreshed by the same during Form_Resize and ClearButton_Click

'blinking LED, if not all labels where allocated
Dim timeaccumLED As Integer  'milliseconds
Const timeblinkLED = 3000 'milliseconds

Private Sub Pause(ByVal milliseconds As Double) 'milliseconds
    Start = Timer * 1000 ' Set start time.
    Do While Timer * 1000 < Start + milliseconds
        DoEvents    ' Yield to other processes.
        If Timer * 1000 < Start Then 'trepassing midnight
          Start = Start - 24! * 60 * 60 * 1000
        End If
    Loop
End Sub

Private Sub AddButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   'remove that dotted ractangle on the button face...
   SuckInFocus.SetFocus
End Sub

Private Sub BlinkLED_Timer()
   Static ff As Boolean
   ff = Not ff
   
   Shape1.Visible = True
   
   Shape1.FillColor = IIf(ff, &HFFFF00, &HC0C000)
   
   timeaccumLED = timeaccumLED + BlinkLED.Interval
   
   If timeaccumLED >= timeblinkLED Then
      Shape1.Visible = False
      BlinkLED.Enabled = False
      timeaccumLED = 0
   End If
End Sub

Private Sub ClearButton_Click()
   If ViewCellsCheck.Value = vbChecked Then
      ret = OccupiedSpaceFun(ChoosenContainer, Label1, ClearContainerRegion + ShowCells, MaxAllocObjsNumber)          'implies showcells
   Else
      ret = OccupiedSpaceFun(ChoosenContainer, Label1, ClearContainerRegion + HideCells, MaxAllocObjsNumber)
   End If
   InitializeButtons
End Sub

Private Sub ClearButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   'remove that dotted ractangle on the button face...
   SuckInFocus.SetFocus
End Sub

Private Sub DesignTimeButton_Click()
   Static ff As Boolean
   ff = Not ff
   DesignTimeLabelSize = ff
   If DesignTimeLabelSize Then
      'do nothing if you want use labels as built at design time
      DesignTimeNameObjects Label1  'this just to name labels (assign their index in caption), while their dimensions remain unchanged
      'assign proper caption to toggling button
      DesignTimeButton.Caption = "design time"
   Else  'change randomly dimensions, at run time
      RunTimeSizeObjects Label1
      'assign proper caption to toggling button
      DesignTimeButton.Caption = "run time"
   End If
 
   ClearButton_Click
End Sub

Private Sub DesignTimeButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   'remove that dotted ractangle on the button face...
   SuckInFocus.SetFocus
End Sub

Private Sub Form_Initialize()
   'backup design time properties of objects
   DesignTimeBackUpRestoreObjects Label1, BackUp
   'change randomly dimensions, at run time (i.e. now)
   RunTimeSizeObjects Label1
   'assign proper caption to toggling button
   DesignTimeButton.Caption = "run time"
   'choose container
   Set ChoosenContainer = Picture1 'OccupiedSpace 'Picture1
End Sub

Private Sub Form_Load()
   SuckInFocus.Left = -SuckInFocus.Left - 2 * SuckInFocus.Width
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button <> 1 Then Exit Sub
   'only if form is the container
   If ChoosenContainer.Name <> "OccupiedSpace" Then Exit Sub
   Static ff As Boolean
   CommandsFrame.Visible = ff
   ff = Not ff

End Sub

Private Sub Form_Resize()
Static NoResize As Boolean
If NoResize Then Exit Sub
  
   With CommandsFrame
      .Move (ScaleWidth - .Width) / 2, ScaleHeight - .Height
   End With
   
   With Picture1
      .Visible = (ChoosenContainer.Name = "Picture1")
      On Error Resume Next
         pixnumber = 3
         marginx = pixnumber * Screen.TwipsPerPixelX
         marginy = pixnumber * Screen.TwipsPerPixelY
         .Move marginx, marginy, ScaleWidth - 2 * marginx, ScaleHeight - 2 * marginy - CommandsFrame.Height
         Err.Clear
      On Error GoTo 0
      RichTextBox1.Move .Left, .Top, .Width, ScaleHeight - .Left '.Height
   End With
   


   'if form is resized too small, then automatic resize to a minimum, so that
   'all controls are shown correctly
   Pause 10  'needed for a smoot transition
'the concept of automatic resizing is this:
'   If UpDown1.Left + 3 * UpDown1.Width > ScaleWidth Then
'      Width = Width * 1.1 'reincrease little by little
'   End If
'   If ScaleHeight < 5 * AddButton.Height Then
'      Height = Height * 1.1
'   End If

   
   'better development (contemporary resizing in width and height):
   If (CommandsFrame.Width > ScaleWidth) Or (LBLwidth > ScaleWidth) Then
      NoResize = True
      Width = Width * 1.1 'reincrease little by little
      If ScaleHeight < 4 * CommandsFrame.Height + LBLheight Then
         NoResize = False
         If ScaleHeight <= 0 Then
            Height = 2 * CommandsFrame.Height
         Else
            Height = Height * 1.1
         End If
      Else
         NoResize = False
         Width = Width * 1.01
      End If
      Exit Sub
   ElseIf ScaleHeight < 4 * CommandsFrame.Height + LBLheight Then
      If ScaleHeight <= 0 Then
         Height = 2 * CommandsFrame.Height
      Else
         Height = Height * 1.1
      End If
      Exit Sub
   End If
   
   
   If ViewCellsCheck.Value = vbChecked Then 'view cells
      ret = OccupiedSpaceFun(ChoosenContainer, Label1, ClearContainerRegion + ShowCells, MaxAllocObjsNumber)
   Else  'hide cells
      ret = OccupiedSpaceFun(ChoosenContainer, Label1, ClearContainerRegion + HideCells, MaxAllocObjsNumber)
   End If

   InitializeButtons  'review this and get proper values from module function
    

End Sub

Private Sub InitializeButtons()
   UpDown1.Max = MaxAllocObjsNumber
   If UpDown1.Max < 1 Then UpDown1.Max = 1
   
   UpDown1.Value = 1
   UpDown1.Min = 0
   AddButtonCaptionRoutine
   
End Sub

Private Sub AddButtonCaptionRoutine()
   plural = IIf(UpDown1.Value = 1, "", "s")
   AddButton.Caption = "add " & Format(UpDown1.Value, "000") & " object" & plural & vbNewLine & "[of " & Format(UpDown1.Max, "000") & " remaining]"
   menucomment.Caption = MaxAllocObjsNumber & " objs random allocation, no collision"
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   'only if form is the container
      Label1(Index).ToolTipText = IIf((ChoosenContainer.Name = "OccupiedSpace"), "click to toggle Commands Frame visibility", "")
End Sub

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button <> 1 Then Exit Sub
   Form_MouseUp Button, 0, 0, 0
End Sub

Private Sub menuabout_Click()
   MsgBox "On ideas of A. Gunnar, email:g68_se@yahoo.se" & _
          vbNewLine & "   Realized by Pietro Cecchi, email: pietrocecchi@inwind.it" & _
          vbNewLine & "August 31st 2K1, posted on www.planet-source-code.com/vb" & _
          vbNewLine & "   Program written in Visual Basic 6" & _
          vbNewLine & _
          vbNewLine & "Did you change the form size?... Have fun :)", vbOKOnly, "      OccupiedSpaceFun, this function allocates randomly many objects (differerent size), no collision!"
End Sub

Private Sub menuusage_Click()
 Static ff As Boolean
 ff = Not ff
 
 If ff Then
   WindowState = vbMaximized
   On Error Resume Next 'if no file
   RichTextBox1.LoadFile App.Path & "\usagenotes.rtf", rtfRTF
   Err.Clear
   On Error GoTo 0
'moved in Form_Resize:
'   With Picture1
'      RichTextBox1.Move .Left, .Top, .Width, ScaleHeight - .Left '.Height
'   End With
   RichTextBox1.ZOrder
   RichTextBox1.Visible = True
   RichTextBox1.BackColor = vbBlue
   RichTextBox1.SelColor = vbYellow
   'temporary tooltiptext on rtb
   RichTextBox1.ToolTipText = "double click to close"
   TemporizeToolTipRTF.Interval = 3000
   TemporizeToolTipRTF.Enabled = True
 Else
'enable this code lineif you want change file contents
'   RichTextBox1.SaveFile App.Path & "\usagenotes.rtf", rtfRTF
   RichTextBox1.Visible = False
 End If

End Sub

Private Sub RichTextBox1_Change()
   RichTextBox1.BackColor = vbBlue
   RichTextBox1.SelColor = vbYellow
End Sub

Private Sub RichTextBox1_DblClick()
   'double click also makes vanish RichText box
   menuusage_Click
End Sub

Private Sub TemporizeToolTipRTF_Timer()
   TemporizeToolTipRTF.Enabled = False
   RichTextBox1.ToolTipText = ""
End Sub

Private Sub ToggleContainer_Click()
   Static ff As Boolean
   ff = Not ff
   Set ChoosenContainer = IIf(ff, OccupiedSpace, Picture1)
   ToggleContainer.Caption = IIf(ff, "container is form", "container is picture")
   Form_Resize
End Sub

Private Sub ToggleContainer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   'remove that dotted ractangle on the button face...
   SuckInFocus.SetFocus
End Sub

Private Sub UpDown1_Change()
   AddButtonCaptionRoutine
   AddButton.Enabled = True
   If UpDown1.Max <= UpDown1.Min Then
      UpDown1.Max = UpDown1.Min
      AddButton.Enabled = False
   End If
   If UpDown1.Value <= 0 Then
      AddButton.Enabled = False
   End If
End Sub

Private Sub AddButton_Click()
   'add UpDown1.Value objects randomly
   CurrentlyShownObjectsNumberSave = CurrentlyShownObjectsNumber
   delta = 0
   
   'call here OccupiedSpaceFun
   Dim allocatenumber As Integer
   allocatenumber = UpDown1.Value

   ret = OccupiedSpaceFun(ChoosenContainer, Label1, AddObjectsToContainer, allocatenumber)

   'if ret>0 then not all objects allocated
   'allocatenumber returns number of correctly allocated objects
   delta = allocatenumber
  
   'stop blinking LED (previous alert)
   BlinkLED.Enabled = False
   timeaccumLED = 0
   DoEvents
   'update controls and globals
   UpDown1.Max = UpDown1.Max - delta
   CurrentlyShownObjectsNumber = CurrentlyShownObjectsNumberSave + delta
   'force change
   UpDown1_Change
   'blink LED if not complete success in allocating
   If ret = ErrorInADD Then 'only some of them or none have been allocated
      'blink LED
      BlinkLED.Interval = 500
      BlinkLED.Enabled = True
   End If

End Sub

Private Sub ViewCellsCheck_Click()
   ReDim VisibleLabels(0 To Label1.Count - 1)
   'ViewCellsCheck.Value= 0 unchecked, 1 checked, 2 grayed
   ViewCellsCheck.Caption = Choose(ViewCellsCheck.Value + 1, "show cells", "hide cells", "hide cells")
   Select Case ViewCellsCheck.Value
      Case vbUnchecked, vbGrayed
         ret = OccupiedSpaceFun(ChoosenContainer, Label1, HideCells)
      Case vbChecked
         ret = OccupiedSpaceFun(ChoosenContainer, Label1, ShowCells)
   End Select
End Sub

Private Sub ViewCellsCheck_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   'remove that dotted ractangle on the button face...
   SuckInFocus.SetFocus
End Sub


