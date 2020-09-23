VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2835
   ClientLeft      =   255
   ClientTop       =   1800
   ClientWidth     =   4770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "WinPaint 1.0"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   $"frmAbout.frx":000C
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Perudow Software"
      BeginProperty Font 
         Name            =   "Script"
         Size            =   24
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
