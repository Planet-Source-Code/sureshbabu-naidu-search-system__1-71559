VERSION 5.00
Begin VB.Form frmReport 
   Caption         =   "Asset Search System"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9315
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   9315
   WindowState     =   2  'Maximized
   Begin VB.Menu mnufilereport 
      Caption         =   "&File"
      Begin VB.Menu mnufilereportdrpt1 
         Caption         =   "&Data Report"
      End
      Begin VB.Menu mnufilereportprpt 
         Caption         =   "&Print Report"
      End
   End
   Begin VB.Menu mnufilereportexit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "frmreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnufilereportdrpt1_Click()
Load DataReport1
DataReport1.Show
End Sub

Private Sub mnufilereportexit_Click()
i = MsgBox("Do you want to exit from User Report ", vbYesNo, "Data Report")
If i = vbYes Then
Unload Me
Unload DataReport1
Else
If i = vbNo Then
frmreport.Show
End If
End If
End Sub

Private Sub mnufilereportprpt_Click()
DataReport1.PrintReport
End Sub
