VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Asset Search System"
   ClientHeight    =   5910
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10215
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   5295
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
            Text            =   "INDIAN"
            TextSave        =   "INDIAN"
            Key             =   "sbr1"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Text            =   "Time"
            TextSave        =   "07:30"
            Key             =   "sbr2"
            Object.ToolTipText     =   "Indian Airlines"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            Text            =   "Date"
            TextSave        =   "14/08/2007"
            Key             =   "sbr3"
            Object.ToolTipText     =   "Indian Airlines"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   2
            Enabled         =   0   'False
            Text            =   "Caps"
            TextSave        =   "CAPS"
            Key             =   "sbr4"
            Object.ToolTipText     =   "Indian Airlines"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   2
            TextSave        =   "NUM"
            Key             =   "sbr5"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   2
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufiledata 
         Caption         =   "&Data Entry And Search"
      End
      Begin VB.Menu mnufilereport 
         Caption         =   "Data &Report"
      End
      Begin VB.Menu mnufilehelp 
         Caption         =   "&Help"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuut 
      Caption         =   "&Utilities"
      Begin VB.Menu mnuutword 
         Caption         =   "&Word"
      End
      Begin VB.Menu mnuutcal 
         Caption         =   "&Calculator"
      End
      Begin VB.Menu mnuutppt 
         Caption         =   "&Power Point"
      End
      Begin VB.Menu mnuutexl 
         Caption         =   "&Excel"
      End
      Begin VB.Menu mnuutie 
         Caption         =   "&Note Pad"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuexit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
str1 = "Indian Airlines"
Load frmADO
frmADO.Show

End Sub

Private Sub mnuabout_Click()
frmADO.Hide
Load frmAbout
frmAbout.Show
End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnufiledata_Click()
Me.Caption = "Data Entry & Search"
Load frmADO
frmADO.Show

End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnufilereport_Click()
Me.Caption = "Data Report"
Load frmreport
frmreport.Show
End Sub


Private Sub mnuutcal_Click()
Dim shellc, fsoc
 Set shellc = CreateObject("WScript.Shell")
 Set fsoc = CreateObject("Scripting.FileSystemObject")
 shellc.Run """C:\WINDOWS\system32\calc.exe"
End Sub

Private Sub mnuutexl_Click()
Dim shelle, fsoe
 Set shelle = CreateObject("WScript.Shell")
 Set fsoe = CreateObject("Scripting.FileSystemObject")
 shelle.Run """C:\Program Files\Microsoft Office\Office\EXCEL.exe"" file.xls"
End Sub

Private Sub mnuutie_Click()
Dim shelln, fsoi
 Set shelln = CreateObject("WScript.Shell")
 Set fson = CreateObject("Scripting.FileSystemObject")
 shelln.Run """C:\WINDOWS\system32\notepad.exe"" file.txt"
End Sub

Private Sub mnuutppt_Click()
Dim shellp, fsop
 Set shellp = CreateObject("WScript.Shell")
 Set fsop = CreateObject("Scripting.FileSystemObject")
 shellp.Run """C:\Program Files\Microsoft Office\Office\POWERPNT.exe"" file.ppt"
End Sub

Private Sub mnuutword_Click()
Dim shell, fso
 Set shell = CreateObject("WScript.Shell")
 Set fso = CreateObject("Scripting.FileSystemObject")
 shell.Run """C:\Program Files\Microsoft Office\Office\WINWORD.exe"" file.doc"
End Sub

Private Sub Timer1_Timer()


'Timer1.Interval = 100
'i = i + 1
'StatusBar1.Panels(1).Text = Left(str1, i)
'If i = Len(str1) Then
'i = 1
'Timer1.Interval = 3000
'End If
End Sub
'End Sub
