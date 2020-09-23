VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login : Asset Search System"
   ClientHeight    =   6270
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   11700
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":08CA
   ScaleHeight     =   3704.523
   ScaleMode       =   0  'User
   ScaleWidth      =   10985.67
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      BackColor       =   &H80000018&
      Height          =   345
      Left            =   1680
      TabIndex        =   0
      Top             =   5040
      Width           =   1605
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H000080FF&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H000080FF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   900
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H80000018&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5640
      Width           =   1605
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "and Chennai. Together with its subsidiary Alliance Air Indian Airlines carrier a total of  over 7.5 million passengers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   240
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   11700
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Indian Airlines flight operations centre around its four main  hubs the main metro cities of Delhi , Mumbai, Calcutta "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   240
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   11670
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "A320, domestic shuttle service, walk in flights and easy fares."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   6315
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "firsts to its credit, including introduction of the wide bodied A300 aircraft on the domestic network,  the fly-by-wire"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   11505
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Indian Airlines has been setting the standards for civil aviation in India since its inception in 1953. It has many"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   11220
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password    :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   5640
      Width           =   1305
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   5040
      Width           =   1320
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by : Ankur, Priyanka, Sundeep and Suresh"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   225
      Left            =   7320
      TabIndex        =   8
      Top             =   6000
      Width           =   4290
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "nationalized to provide well coordinated, adequate, safe, efficient and economical air services."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   9810
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Indian Airlines was given the task to assimilate various dimensions of the eight private airlines, which were"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   10950
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "responsibility of providing air transportation within the country as well as to the neighbouring countries."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   10530
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Indian Airlines came into being with the enactment of the Air Corporations Act 1953 and was entrusted with the "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   11265
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Unload Me
    End Sub

Private Sub cmdOK_Click()
    'check for correct password
    If txtUserName = "indian" And txtPassword = "indian" Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        Unload Me
        Load MDIForm1
        MDIForm1.Show
    Else
        MsgBox "Invalid UserName or Password, try again!", , "Login"
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub



