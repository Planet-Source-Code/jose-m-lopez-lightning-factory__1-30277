VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmControls 
   Caption         =   "Controls"
   ClientHeight    =   7575
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4890
   LinkTopic       =   "Form2"
   ScaleHeight     =   7575
   ScaleWidth      =   4890
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H0080C0FF&
      Caption         =   "Stop"
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   6840
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame framePlay 
      Caption         =   "Playlist"
      Height          =   6615
      Left            =   3000
      TabIndex        =   65
      Top             =   0
      Width           =   1815
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   120
         TabIndex        =   69
         Top             =   5640
         Width           =   495
         Begin VB.CommandButton cmdReplace 
            BackColor       =   &H00C0C0FF&
            Height          =   195
            Left            =   240
            Picture         =   "frmControls.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   80
            ToolTipText     =   "Replace highlighted item with parameters"
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton cmdClearList 
            BackColor       =   &H00C0C0FF&
            Height          =   195
            Left            =   240
            Picture         =   "frmControls.frx":0496
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Clear playlist"
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00C0C0FF&
            Height          =   195
            Left            =   0
            Picture         =   "frmControls.frx":0910
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Add parameters to end of playlist"
            Top             =   120
            Width           =   255
         End
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H00C0C0FF&
            Height          =   195
            Left            =   0
            Picture         =   "frmControls.frx":0D8A
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Delete highlighted item"
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton cmdRetrieve 
            BackColor       =   &H00C0C0FF&
            Height          =   195
            Left            =   0
            Picture         =   "frmControls.frx":1204
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Retrieve highlighted item to parameters; Also reteive by dblclkng playlist"
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton cmdInsert 
            BackColor       =   &H00C0C0FF&
            Height          =   195
            Left            =   240
            Picture         =   "frmControls.frx":169A
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Insert parameters above highlighted item"
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Left            =   720
         TabIndex        =   67
         Top             =   5640
         Width           =   975
         Begin VB.OptionButton optHighLightExe 
            Caption         =   "HLon"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   83
            ToolTipText     =   "highlight item as executing (slower!)"
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton optHighLightExe 
            Caption         =   "HLoff"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   82
            ToolTipText     =   "don't highlight item as executing (faster)"
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton cmdPlayOnly 
            BackColor       =   &H00C0FFFF&
            Height          =   195
            Left            =   720
            Picture         =   "frmControls.frx":1B14
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Play highlighted item only"
            Top             =   120
            Width           =   255
         End
         Begin VB.CommandButton cmdPlayFrom 
            BackColor       =   &H00C0FFFF&
            Height          =   195
            Left            =   480
            Picture         =   "frmControls.frx":1FAA
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Play from highlighted item"
            Top             =   120
            Width           =   255
         End
         Begin VB.CommandButton cmdPlayAll 
            BackColor       =   &H00C0FFFF&
            Height          =   195
            Left            =   240
            Picture         =   "frmControls.frx":2440
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Play everything"
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.ListBox LstMain 
         BackColor       =   &H00C0FFFF&
         Height          =   5325
         Left            =   120
         TabIndex        =   66
         ToolTipText     =   "Playlist (double click to retrieve item))"
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame frameControls 
      Caption         =   "Parameters"
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   18
         Left            =   120
         TabIndex        =   77
         Top             =   6600
         Width           =   2535
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   22
            Left            =   0
            TabIndex        =   78
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblParameters 
            Caption         =   "delay btwn displays"
            Height          =   255
            Index           =   18
            Left            =   840
            TabIndex        =   79
            ToolTipText     =   "delay in milliseconds between each display"
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   62
         Top             =   480
         Width           =   1335
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   1
            Left            =   0
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblParameters 
            Caption         =   "y start"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   63
            ToolTipText     =   "where ray will start"
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Text            =   "Text1"
         ToolTipText     =   "title"
         Top             =   7080
         Width           =   2655
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   17
         Left            =   120
         TabIndex        =   60
         Top             =   6240
         Width           =   2535
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   21
            Left            =   0
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblParameters 
            Caption         =   "delay btwn segments"
            Height          =   255
            Index           =   17
            Left            =   840
            TabIndex        =   61
            ToolTipText     =   "delay in milliseconds between ray segments"
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   16
         Left            =   120
         TabIndex        =   58
         Top             =   5880
         Width           =   2535
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   20
            Left            =   0
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblParameters 
            Caption         =   "delay btwn rays"
            Height          =   255
            Index           =   16
            Left            =   840
            TabIndex        =   59
            ToolTipText     =   "delay in milliseconds between rays"
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   15
         Left            =   120
         TabIndex        =   56
         Top             =   5520
         Width           =   2655
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   19
            Left            =   0
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblParameters 
            Caption         =   "# displays (minus = inf)"
            Height          =   255
            Index           =   15
            Left            =   840
            TabIndex        =   57
            ToolTipText     =   "quantity of display to show, neg loops always"
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Height          =   375
         Left            =   1680
         TabIndex        =   54
         Top             =   240
         Width           =   495
         Begin VB.CommandButton cmdStart 
            BackColor       =   &H00C0FFC0&
            Height          =   195
            Left            =   240
            Picture         =   "frmControls.frx":28BA
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "execute parameter text boxes"
            Top             =   120
            Width           =   255
         End
         Begin VB.CommandButton cmdClear 
            BackColor       =   &H00C0FFC0&
            Height          =   195
            Left            =   0
            Picture         =   "frmControls.frx":2D34
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "clear the screen"
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   52
         Top             =   5160
         Width           =   2415
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   18
            Left            =   960
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   17
            Left            =   480
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0C0FF&
            Height          =   285
            Index           =   16
            Left            =   0
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblParameters 
            Caption         =   "outer color"
            Height          =   255
            Index           =   14
            Left            =   1560
            TabIndex        =   53
            ToolTipText     =   "color outside of radius"
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   50
         Top             =   4800
         Width           =   2415
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   15
            Left            =   960
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   14
            Left            =   480
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0C0FF&
            Height          =   285
            Index           =   13
            Left            =   0
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblParameters 
            Caption         =   "inner color"
            Height          =   255
            Index           =   13
            Left            =   1560
            TabIndex        =   51
            ToolTipText     =   "color inside of radius"
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   46
         Top             =   4440
         Width           =   2535
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   12
            Left            =   0
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblParameters 
            Caption         =   "raywdth (0 to -15=cls)"
            Height          =   255
            Index           =   12
            Left            =   840
            TabIndex        =   47
            ToolTipText     =   "ray width, values between 0 and -15 will clear to qbcolor"
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   44
         Top             =   4080
         Width           =   1935
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   11
            Left            =   0
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblParameters 
            Caption         =   "y/2 centering"
            Height          =   255
            Index           =   11
            Left            =   840
            TabIndex        =   45
            ToolTipText     =   "for centering y rnd dev. (normally y rnd dev/2)"
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   42
         Top             =   3720
         Width           =   1935
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   10
            Left            =   0
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblParameters 
            Caption         =   "x/2 centering"
            Height          =   255
            Index           =   10
            Left            =   840
            TabIndex        =   43
            ToolTipText     =   "for centering x rnd dev. (normally x rnd dev/2)"
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   40
         Top             =   3360
         Width           =   1695
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   9
            Left            =   0
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblParameters 
            Caption         =   "y rnd dev"
            Height          =   255
            Index           =   9
            Left            =   840
            TabIndex        =   41
            ToolTipText     =   "random deviation to be added to linear y axis increment of ray"
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   38
         Top             =   3000
         Width           =   1575
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   8
            Left            =   0
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblParameters 
            Caption         =   "x rnd dev"
            Height          =   255
            Index           =   8
            Left            =   840
            TabIndex        =   39
            ToolTipText     =   "random deviation to be added to linear x axis increment of ray"
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   36
         Top             =   2640
         Width           =   1455
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   7
            Left            =   0
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblParameters 
            Caption         =   "radius"
            Height          =   255
            Index           =   7
            Left            =   840
            TabIndex        =   37
            ToolTipText     =   "determines where inner and outer colors divide"
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   34
         Top             =   2280
         Width           =   2655
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   6
            Left            =   0
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblParameters 
            Caption         =   "# fades (0 = no fades)"
            Height          =   255
            Index           =   6
            Left            =   840
            TabIndex        =   35
            ToolTipText     =   "number of ray fading steps. 0 & neg = no fades"
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   32
         Top             =   1920
         Width           =   2415
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   5
            Left            =   0
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblParameters 
            Caption         =   "segments per ray"
            Height          =   255
            Index           =   5
            Left            =   840
            TabIndex        =   33
            ToolTipText     =   "quantity segments/points in each ray"
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   1560
         Width           =   2535
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   4
            Left            =   0
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblParameters 
            Caption         =   "rays per display"
            Height          =   255
            Index           =   4
            Left            =   840
            TabIndex        =   31
            ToolTipText     =   "quantity lightning rays to show in one display"
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   2535
         Begin VB.OptionButton optAngle 
            Caption         =   "Option1"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   49
            Top             =   120
            Width           =   255
         End
         Begin VB.OptionButton optAngle 
            Caption         =   "Option1"
            Height          =   255
            Index           =   0
            Left            =   1920
            TabIndex        =   48
            Top             =   120
            Width           =   255
         End
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   3
            Left            =   0
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblParameters 
            Caption         =   "upper angle"
            Height          =   255
            Index           =   3
            Left            =   840
            TabIndex        =   29
            ToolTipText     =   "section of ""imaginary pie chart"" where rays will display randomly"
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   2655
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   2
            Left            =   0
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   " off   on"
            Height          =   255
            Left            =   1920
            TabIndex        =   81
            ToolTipText     =   "pie chart off/on"
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblParameters 
            Caption         =   "lower angle"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   27
            ToolTipText     =   "section of ""imaginary pie chart"" where rays will display randomly"
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame frmParameters 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1335
         Begin VB.TextBox txtParameters 
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblParameters 
            Caption         =   "x start"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   25
            ToolTipText     =   "where ray will start"
            Top             =   0
            Width           =   495
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileOpenWithAppend 
         Caption         =   "Append"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnFileSaveAs 
         Caption         =   "Save as"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
'LstMain.AddItem string            append item string
'LstMain.AddItem string, where     insert item string, 0 = 1st line
'LstMain.List(0) = string          replace item string
'string = LstMain.List(0)          return item string
'int = LstMain.ListCount           return number of items
'string = LstMain.List(LstMain.ListIndex) return highlighted item string
'int = LstMain.ListIndex           return highlighted item#, 0 to last, -1 = nothing highlighted
'LstMain.ListIndex = 2             highlight item #2
'LstMain.ListIndex = -1            dehighlight all
'**************************************************
Dim strTextBoxesToOneLine As String
Dim strOneLineToTextBoxes As String

Private Sub cmdStop_Click()
If booStopped = True Then Exit Sub
booStop = True
End Sub
Private Sub cmdStart_Click()
'run plasma routine according to textboxes
rtnStart
End Sub
Private Sub rtnStart()
frameControls.Enabled = False
framePlay.Enabled = False
booStopped = False
frmScreen.Show
rtnPlasma1
frameControls.Enabled = True
framePlay.Enabled = True
booStopped = True
booStop = False
frmControls.Show
End Sub

Private Sub cmdClearList_Click()
'Clear Main Listbox
Dim intMsg As Integer
intMsg = MsgBox("This will clear listbox", vbOKCancel, "Clear")
If intMsg = vbCancel Then Exit Sub
LstMain.Clear
strFileName = "" 'To prevent writing over on mnuFileSave
End Sub

Private Sub cmdClear_Click()
'clear screen
frmScreen.Cls
End Sub


Private Sub cmdAdd_Click()
'Add entry to Main ListBox
rtnTextBoxesToOneLine 'get all entries
LstMain.AddItem strTextBoxesToOneLine 'add entries to Main ListBox
LstMain.ListIndex = LstMain.ListCount - 1
End Sub

Private Sub cmdInsert_Click()
'Insert entry to Main ListBox above Hightlighted item
Dim X As Integer
If LstMain.ListIndex = -1 Then Exit Sub 'if no item selected then leave
LstMain.AddItem "temporary dummy"
    For X = LstMain.ListCount - 1 To LstMain.ListIndex + 1 Step -1
    LstMain.List(X) = LstMain.List(X - 1)
    Next X
rtnTextBoxesToOneLine
LstMain.List(LstMain.ListIndex) = strTextBoxesToOneLine
End Sub

Private Sub cmdReplace_Click()
'Replace Hightlighted item in Main ListBox
If LstMain.ListIndex = -1 Then Exit Sub 'if no item selected then leave
rtnTextBoxesToOneLine
LstMain.List(LstMain.ListIndex) = strTextBoxesToOneLine
End Sub

Private Sub cmdRetrieve_Click()
'Get highlighted listbox item and decode to textboxes
If LstMain.ListIndex = -1 Then Exit Sub 'if no item selected then leave
strOneLineToTextBoxes = LstMain.List(LstMain.ListIndex) 'item to retrieve
rtnOneLineToTextBoxes 'decode to textboxes
End Sub

Private Sub LstMain_DblClick()
'Get highlighted listbox item and decode to textboxes
If LstMain.ListIndex = -1 Then Exit Sub 'if no item selected then leave
strOneLineToTextBoxes = LstMain.List(LstMain.ListIndex) 'item to retrieve
rtnOneLineToTextBoxes 'decode to textboxes
End Sub

Private Sub cmdDelete_Click()
'Delete highlighted item
Dim X As Integer
Dim txtArray() As String
Dim intListCount As Integer
Dim intReHighLight As Integer
ReDim txtArray(LstMain.ListCount)
intListCount = LstMain.ListCount

If LstMain.ListIndex = -1 Then Exit Sub 'if no item selected then leave
intReHighLight = LstMain.ListIndex
    'since an item is being removed simply move each item
    'after it one step up overwriting each other
    For X = LstMain.ListIndex To LstMain.ListCount - 1
    LstMain.List(X) = LstMain.List(X + 1)
    Next X

    'Now we need to remove the extra item (last entry in listbox)
    'we do this by temporarily saving listbox to array, clearing listbox,
    'and then reloading listbox with entries
    For X = 0 To intListCount - 2
    txtArray(X) = LstMain.List(X)
    Next X
    
    LstMain.Clear
    For X = 0 To intListCount - 2
    LstMain.List(X) = txtArray(X)
    Next X
If intReHighLight > LstMain.ListCount - 1 Then
LstMain.ListIndex = LstMain.ListCount - 1
Else
LstMain.ListIndex = intReHighLight
End If
End Sub

Private Sub rtnPrint()
If frmScreen.BackColor = vbWhite Then
frmScreen.ForeColor = vbBlack
Else
frmScreen.ForeColor = vbWhite
End If
frmScreen.CurrentX = 50
frmScreen.CurrentY = 0
frmScreen.Print "Click=Stop DblClk=End   " & frmControls.txtTitle.Text
End Sub
Private Sub cmdPlayAll_Click()
'Run all entries in listbox
Dim X As Integer
If LstMain.ListCount <= 0 Then Exit Sub
frameControls.Enabled = False  'disable user controls so no funny business while busy
framePlay.Enabled = False      'disable user controls so no funny business while busy
booStop = False
booStopped = False
frmScreen.Show
If booHighlightOn = True Then frmControls.Show 'if option to highlight execution selected then show this form too
    'Execute each playlist item one by one
    For X = 0 To LstMain.ListCount - 1 'run each playlist item
    strOneLineToTextBoxes = LstMain.List(X) 'get an item from playlist
    rtnOneLineToTextBoxes 'load it in the textboxes
    If booHighlightOn = True Then LstMain.ListIndex = X 'if option to highlight execution selected then highlight here
    rtnPrint
    rtnPlasma1 'now display according to the textbox parameters
    DoEvents
    If booStop = True Then Exit For 'if user pressed stop then stop
    Sleep txtParameters(22) 'delay between displays
    Next X
frameControls.Enabled = True 'finished so renable user controls
framePlay.Enabled = True     'finished so renable user controls
booStop = False
booStopped = True
frmControls.Show
End Sub
Private Sub cmdPlayFrom_Click()
'Run entries from highlight to end of listbox
Dim X As Integer
If LstMain.ListCount <= 0 Then Exit Sub 'nothing in listbox
If LstMain.ListIndex = -1 Then Exit Sub 'nothing selected
frameControls.Enabled = False  'disable user controls so no funny business while busy
framePlay.Enabled = False      'disable user controls so no funny business while busy
booStop = False
booStopped = False
frmScreen.Show
If booHighlightOn = True Then frmControls.Show 'if option to highlight execution selected then show this form too
    'Execute each playlist item one by one
    For X = LstMain.ListIndex To LstMain.ListCount - 1 'run each playlist item
    strOneLineToTextBoxes = LstMain.List(X) 'get an item from playlist
    rtnOneLineToTextBoxes 'load it in the textboxes
    If booHighlightOn = True Then LstMain.ListIndex = X 'if option to highlight execution selected then highlight here
    rtnPrint
    rtnPlasma1 'now display according to the textbox parameters
    DoEvents
    If booStop = True Then Exit For 'if user pressed stop then stop
    Sleep txtParameters(22)
    Next X
frameControls.Enabled = True 'finished so renable user controls
framePlay.Enabled = True     'finished so renable user controls
booStop = False
booStopped = True
frmControls.Show
End Sub
Private Sub cmdPlayOnly_Click()
'Run highlighted entry in listbox
Dim X As Integer
If LstMain.ListCount <= 0 Then Exit Sub 'nothing in listbox
If LstMain.ListIndex = -1 Then Exit Sub 'nothing selected

frameControls.Enabled = False  'disable user controls so no funny business while busy
framePlay.Enabled = False      'disable user controls so no funny business while busy
booStop = False
booStopped = False
frmScreen.Show
    strOneLineToTextBoxes = LstMain.List(LstMain.ListIndex) 'get highlighted item from listbox
    rtnOneLineToTextBoxes 'load it in the textboxes
    rtnPrint
    rtnPlasma1 'now display according to the textbox paremeters
frameControls.Enabled = True 'finished so renable user controls
framePlay.Enabled = True     'finished so renable user controls
booStop = False
booStopped = True
frmControls.Show
End Sub

Private Sub mnuFileExit_Click()
'end
Dim intMsg As Integer
intMsg = MsgBox("This will end program, continue?", vbOKCancel, "Caution")
If intMsg = vbCancel Then Exit Sub
Unload frmControls
Unload frmScreen
End
End Sub

Private Sub mnuFileOpen_Click()
'Open a Plasma Data File and place it on Main Listbox
Dim tmpText As String
                           'Common Dialog Window
'***************************************************************************
CommonDialog1.CancelError = True  'Enable on error or cancel GoTo
On Error GoTo cancelPressed
CommonDialog1.DefaultExt = ".dat" 'in case user doesn't type extension
CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
CommonDialog1.DialogTitle = "Open Plasma Data File"        'Title displayed
CommonDialog1.InitDir = App.Path                'Start Directory
'Format     object.Filter [= description1 |filter1 |description2 |filter2...]
'Example               "Text (*.txt)|*.txt| Pictures (*.bmp;*.ico)|*.bmp;*.ico"
CommonDialog1.Filter = "Plasma File (*.dat)|*.dat"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
strFileName = CommonDialog1.FileName
frmControls.Caption = strFileName
'***************************************************************************
    '*******************************************
    'decode file to Main Listbox
    LstMain.Clear
    Open strFileName For Input As #1
    Do While Not EOF(1) 'check for end of file.
    Line Input #1, tmpText
    LstMain.AddItem tmpText
    Loop
    '*******************************************
cancelPressed:
Close #1   ' Close file.
End Sub

Private Sub mnuFileOpenWithAppend_Click()
'Open a Plasma Data File and append it to Main Listbox
Dim tmpText As String
Dim tmpFileName As String
                           'Common Dialog Window
'***************************************************************************
CommonDialog1.CancelError = True  'Enable on error or cancel GoTo
On Error GoTo cancelPressed
CommonDialog1.DefaultExt = ".dat" 'in case user doesn't type extension
CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
CommonDialog1.DialogTitle = "Append"        'Title displayed
CommonDialog1.InitDir = App.Path                'Start Directory
'Format     object.Filter [= description1 |filter1 |description2 |filter2...]
'Example               "Text (*.txt)|*.txt| Pictures (*.bmp;*.ico)|*.bmp;*.ico"
CommonDialog1.Filter = "Plasma File (*.dat)|*.dat"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
tmpFileName = CommonDialog1.FileName
'***************************************************************************
    '*******************************************
    'decode and append file to Main Listbox
    Open tmpFileName For Input As #1
    Do While Not EOF(1) 'check for end of file.
    Line Input #1, tmpText
    LstMain.AddItem tmpText
    Loop
    '*******************************************
cancelPressed:
Close #1   ' Close file.
End Sub

Private Sub mnuFileSave_Click()
'Save the Main Listbox
Dim X As Integer
    If LstMain.ListCount <= 0 Then
    X = MsgBox("Nothing to save!", vbOKOnly, "oops")
    Exit Sub
    End If
    
    If strFileName = "" Then
    rtnSaveAs
    Exit Sub
    End If

Dim intMsg As Integer
intMsg = MsgBox("This will overwrite " & strFileName & " , continue?", vbOKCancel, "Caution")
If intMsg = vbCancel Then Exit Sub
    
    frmControls.Caption = strFileName
    Open strFileName For Output As #1
    For X = 0 To LstMain.ListCount - 1
    Print #1, LstMain.List(X)
    Next X
    Close #1
End Sub

Private Sub mnFileSaveAs_Click()
'Save the Main Listbox using the Common Dialog Window
rtnSaveAs
End Sub

Private Sub rtnSaveAs()
'Save the Main Listbox using the Common Dialog Window
Dim tmpText As String
Dim X As Integer
    If LstMain.ListCount <= 0 Then
    X = MsgBox("Nothing to save!", vbOKOnly, "oops")
    Exit Sub
    End If
                         'Common Dialog Window
'***************************************************************************
CommonDialog1.CancelError = True
  On Error GoTo cancelPressed
CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt
CommonDialog1.DialogTitle = "Save Plasma Data File"
CommonDialog1.InitDir = App.Path
CommonDialog1.Filter = "Plasma File (*.dat)|*.dat"
CommonDialog1.FileName = strFileName 'puts in dialog filname
CommonDialog1.ShowSave
strFileName = CommonDialog1.FileName
frmControls.Caption = strFileName
'***************************************************************************
Open strFileName For Output As #1
    For X = 0 To LstMain.ListCount - 1
    Print #1, LstMain.List(X)
    Next X
Close #1
cancelPressed:

End Sub

Private Sub optAngle_Click(Index As Integer)
'Option for selecting whether or not you want the program to operate
'in a section of a circle, like a pie chart
Select Case Index
    Case 0: intbooAngleInc = 0 'disable pie chart type angle direction
    Case 1: intbooAngleInc = 1 ' enable pie chart type angle direction
End Select
End Sub

Private Sub rtnTextBoxesToOneLine()
'Enter with Textboxes containing entries
'Leave with strTextBoxesToOneLine containing all entries
Dim X As Integer
'get title
strTextBoxesToOneLine = frmControls.txtTitle.Text & ","
    'get all other textboxes from array
    For X = 0 To 22
    strTextBoxesToOneLine = strTextBoxesToOneLine & frmControls.txtParameters(X).Text & ","
    Next X
'get option selected
If frmControls.optAngle(0).Value = True Then
strTextBoxesToOneLine = strTextBoxesToOneLine & "1"
Else
strTextBoxesToOneLine = strTextBoxesToOneLine & "0"
End If
End Sub

Private Sub rtnOneLineToTextBoxes()
'Enter with strOneLineToTextBoxes holding one line of entries
'Leave with this information decoded to all the textboxes
Dim X As Integer, Y As Integer
Dim intListLineCount As Integer
Dim intCount As Integer
Dim intCountHold As Integer
Dim intLetterCount As Integer

intLetterCount = 0
    'extract title to textbox
    For X = 1 To InStr(1, strOneLineToTextBoxes, ",", 1)
    intLetterCount = intLetterCount + 1
    Next X
frmControls.txtTitle.Text = Mid$(strOneLineToTextBoxes, 1, intLetterCount - 1)
intCountHold = X 'one past last comma

    'extract all other parameters to textboxes
    For Y = 0 To 22
    intLetterCount = 0
        For X = intCountHold To InStr(intCountHold, strOneLineToTextBoxes, ",", 1)
        intLetterCount = intLetterCount + 1
        Next X
        frmControls.txtParameters(Y).Text = Mid$(strOneLineToTextBoxes, intCountHold, intLetterCount - 1)
        intCountHold = X 'one past last comma
    Next Y

    'extract option to select
    If Mid$(strOneLineToTextBoxes, intCountHold, Len(strOneLineToTextBoxes) - (intCountHold - 1)) = "1" Then
    frmControls.optAngle(0).Value = True
    Else
    frmControls.optAngle(1).Value = True
    End If

End Sub

Private Sub optHighLightExe_Click(Index As Integer)
Select Case Index
    Case 0: booHighlightOn = False
    Case 1: booHighlightOn = True
End Select
End Sub

Private Sub txtParameters_KeyPress(Index As Integer, KeyAscii As Integer)
'Run the Plasma routine if the enter key is pressed
'while entering data in any of the textboxes
Dim char As String
char = Chr(KeyAscii)
    If char = Chr(13) Then ' Detect enter pressed
    rtnStart
    booStopped = True
    Exit Sub
    End If
End Sub

Private Sub txtTitle_GotFocus()
'Highlight textbox
txtTitle.SelStart = 0
txtTitle.SelLength = Len(txtTitle)
End Sub

Private Sub txtParameters_GotFocus(Index As Integer)
'Highlight textbox
txtParameters(Index).SelStart = 0
txtParameters(Index).SelLength = Len(txtParameters(Index))
End Sub


