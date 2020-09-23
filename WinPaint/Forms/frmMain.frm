VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmMain 
   BackColor       =   &H00C00000&
   Caption         =   "WinPaint 1.0"
   ClientHeight    =   6195
   ClientLeft      =   4005
   ClientTop       =   2550
   ClientWidth     =   8535
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6195
   ScaleWidth      =   8535
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   8535
      TabIndex        =   13
      Top             =   5940
      Width           =   8535
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "(c) 2007 By Perudow Software"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Tools"
      ForeColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C00000&
         Caption         =   "Off"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   2880
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C00000&
         Caption         =   "On"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Take Snapshot"
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save Snapshot"
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Width:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Snapshot Size::"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Stretch:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   2520
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5520
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Snapshot Preview - "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "WinPaint 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   3855
      Left            =   3240
      Top             =   1080
      Width           =   5055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFsave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuftss 
         Caption         =   "Take Snapshot"
      End
   End
   Begin VB.Menu mnhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error GoTo Error
    'Common dialog save routine
    With CD1
        .DialogTitle = "Save Capture To..."
        .Filter = "JPEG File (*.jpg)|*.jpg"
        .CancelError = True
        .Flags = &H2  'Overwrite prompt
        .ShowSave
        If .FileName = "" Then GoTo Error
        'Save capture as bitmap
        SavePicture Image1.Picture, .FileName
    End With
    
Error:   'Canceled
End Sub

Private Sub Command2_Click()
'Hide form so it isn't shown in picture too
Me.Hide
'Capture Screen

snap1 Me
'Show form
Me.Show
'Set window to maximized
'Me.WindowState = 2
End Sub


Sub Form_Resize()
Image1.Left = 3240
'Image1.Width = Me.Width - 3500
'Image1.Height = Me.Height - 2000
Image1.Width = Text2.Text
Image1.Height = Text1.Text
Frame1.Height = Me.Height - 1500
Text1.Text = Me.Height - 2300
Text2.Text = Me.Width - 3500

End Sub
Private Sub Form_Load()
'Show Message Box
Text1.Text = Me.Height - 2300
Text2.Text = Me.Width - 3500
Option1.Value = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
'End Program
End
End Sub

Private Sub mnuabout_Click()
frmAbout.Show
End Sub

Private Sub Option1_Click()
Image1.Stretch = True
Label2.Caption = "Snapshot Preview - Stretch: On"
Form_Resize
End Sub

Private Sub Option2_Click()
Image1.Stretch = False
Label2.Caption = "Snapshot Preview - Stretch: Off"
Form_Resize
End Sub
Private Sub text1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Asc(vbCr)
KeyAscii = 0
Case 8, 46
Case 47 To 58
Case Else
KeyAscii = 0
End Select
End Sub
Private Sub text2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Asc(vbCr)
KeyAscii = 0
Case 8, 46
Case 47 To 58
Case Else
KeyAscii = 0
End Select
End Sub
Private Sub Text1_Change()
Image1.Height = Text1.Text
If Text1.Text < "0" Then
Text1.Text = "0"
End If
End Sub

Private Sub Text2_Change()
Image1.Width = Text2.Text
End Sub
