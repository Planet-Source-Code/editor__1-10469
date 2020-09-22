VERSION 5.00
Begin VB.Form frmHelp 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editor Help"
   ClientHeight    =   5265
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Editor Help"
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5775
      Begin VB.Label Label6 
         Caption         =   $"frmHelp.frx":0742
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   3600
         Width           =   5535
      End
      Begin VB.Label Label5 
         Caption         =   $"frmHelp.frx":07E4
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   5535
      End
      Begin VB.Label Label4 
         Caption         =   $"frmHelp.frx":08C5
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Width           =   5535
      End
      Begin VB.Label Label3 
         Caption         =   $"frmHelp.frx":0957
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Watch the status bar on the bottom of the program screen.  It tells you if your file has changed, been saved, or newly loaded."
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2010
         Left            =   3360
         Picture         =   "frmHelp.frx":09E1
         ToolTipText     =   "This is me, hard at work.."
         Top             =   240
         Width           =   2310
      End
      Begin VB.Label Label1 
         Caption         =   "I can't imagine that you're needing much help using this program, but in case you do, here are some tips:"
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   120
      Picture         =   "frmHelp.frx":289C
      ToolTipText     =   "http://members.home.com/fordpref"
      Top             =   4440
      Width           =   4425
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
    Unload Me
End Sub
