VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmADO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asset Search System"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   10695
   Icon            =   "FRMADO~1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Palette         =   "FRMADO~1.frx":08CA
   Picture         =   "FRMADO~1.frx":4D5A
   ScaleHeight     =   8790
   ScaleWidth      =   10695
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3720
      Top             =   2640
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\try\assets.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\try\assets.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from asset_table"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   6600
      Top             =   4440
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   8640
      ScaleHeight     =   615
      ScaleWidth      =   6615
      TabIndex        =   39
      Top             =   5040
      Width           =   6615
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   8880
      Picture         =   "FRMADO~1.frx":91EA
      ScaleHeight     =   5295
      ScaleWidth      =   6375
      TabIndex        =   38
      Top             =   0
      Width           =   6375
   End
   Begin VB.PictureBox picStatBoxReg 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   8655
      TabIndex        =   26
      Top             =   5040
      Width           =   8655
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Last"
         Height          =   350
         Left            =   6480
         MaskColor       =   &H008080FF&
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   100
         UseMaskColor    =   -1  'True
         Width           =   1425
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Next >>"
         Height          =   350
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1305
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFC0&
         Caption         =   "<< Prev"
         Height          =   350
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1425
      End
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFC0&
         Caption         =   "First"
         Height          =   350
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1545
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FRMADO~1.frx":D67A
      Height          =   5415
      Left            =   0
      TabIndex        =   25
      Top             =   5640
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   9551
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   -2147483630
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "DATE"
         Caption         =   "DATE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "PARTICULARS"
         Caption         =   "PARTICULARS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "ASSET"
         Caption         =   "ASSET"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "GRN"
         Caption         =   "GRN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "JV_IDN_CV"
         Caption         =   "JV_IDN_CV"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "SUPPLIER"
         Caption         =   "SUPPLIER"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "DEBIT"
         Caption         =   "DEBIT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "CREDIT"
         Caption         =   "CREDIT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000018&
      Caption         =   "&Read Input"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000018&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindNext 
      BackColor       =   &H80000018&
      Caption         =   "Find &Next"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindFirst 
      BackColor       =   &H80000018&
      Caption         =   "Find &First"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboField 
      Height          =   315
      ItemData        =   "FRMADO~1.frx":D68F
      Left            =   1440
      List            =   "FRMADO~1.frx":D691
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   4200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text9 
      DataField       =   "CREDIT"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   18
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      DataField       =   "DEBIT"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   17
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      DataField       =   "SUPPLIER"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      DataField       =   "JV_IDN_CV"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      DataField       =   "GRN"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   14
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      DataField       =   "ASSET"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   1920
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      DataField       =   "PARTICULARS"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      DataField       =   "DATE"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "##/##/##"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Top             =   1200
      Width           =   1695
   End
   Begin VB.PictureBox picButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5145
      Left            =   6600
      ScaleHeight     =   5145
      ScaleWidth      =   1995
      TabIndex        =   4
      Top             =   0
      Width           =   1995
      Begin VB.CommandButton Command4 
         BackColor       =   &H000000C0&
         Caption         =   "List All"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H000000C0&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H000000C0&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H000000C0&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H000000C0&
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdAbout 
         BackColor       =   &H000000C0&
         Caption         =   "A&bout"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H000000C0&
         Caption         =   "Fi&nd..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000C0&
         Caption         =   "&Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   5
         Height          =   4815
         Left            =   240
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Basis:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit (Amt):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Debit (Amt):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblAngka 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label lblField 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "JV/DN/CV"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "GRN:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Asset:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Particulars:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   615
   End
End
Attribute VB_Name = "frmADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag, update
Dim a_name, grn, jv, amt_dr, amt_cr, choice, sup As String
'Dim dat as string


Private Sub cmdAbout_Click()
frmADO.Hide
Load frmAbout
frmAbout.Show
  
End Sub

Private Sub cmdAdd_Click()

CheckNavigation
update = 0


cmdAdd.Enabled = False
cmdCancel.Enabled = True
Command1.Enabled = True
cmdDelete.Enabled = False
cmdEdit.Enabled = False
cmdFind.Enabled = False

     If ((Adodc1.Recordset.EOF = True) And (Adodc1.Recordset.BOF = True)) Then
     MsgBox ("Database is Empty!!!")
     Adodc1.Recordset.AddNew
          
   Else

Adodc1.Recordset.MoveLast
Adodc1.Recordset.AddNew

Text2.SetFocus
Text2.Text = Date
End If
End Sub

Private Sub cmdCancel_Click()
CheckNavigation
'On Error Resume Next
If update = 0 Then
'MsgBox Adodc1.Recordset.Fields(0)
'MsgBox Adodc1.Recordset.Fields(2)

Adodc1.Recordset.CancelBatch
Adodc1.Refresh
'Adodc1.Recordset.Delete
End If

cmdAdd.Enabled = True
cmdCancel.Enabled = False
Command1.Enabled = False
cmdDelete.Enabled = True
cmdEdit.Enabled = True
cmdFind.Enabled = True

'qqqqqqqqqqqqqqqqqqqqqqqqqqqqcmdClose.Enabled = True


End Sub



Private Sub cmdDelete_Click()

'MsgBox Adodc1.Recordset.Fields(0)
choice = MsgBox("Are You Sure You Want To Delete This Record ??", vbYesNo + vbExclamation, "Delete?")
',buttons as vbmsgboxstyle= vbOKCancel")

If choice = vbYes Then
     If ((Adodc1.Recordset.EOF = True) And (Adodc1.Recordset.BOF = True)) Then
     MsgBox ("Database is Empty!!!")
   Else
   Adodc1.Recordset.Delete
    'Adodc1.Recordset.MovePrevious
   MsgBox ("Record Sucessfuly Deleted!!!")
    
End If
End If
CheckNavigation
End Sub

Private Sub cmdEdit_Click()
Adodc1.Refresh

If ((Adodc1.Recordset.EOF = True) And (Adodc1.Recordset.BOF = True)) Then
     MsgBox ("Database is Empty!!!")
   Else
   
update = 0
cmdAdd.Enabled = False
cmdCancel.Enabled = True
Command1.Enabled = True
cmdDelete.Enabled = False
cmdEdit.Enabled = False
cmdFind.Enabled = False

CheckNavigation
Adodc1.Recordset.update
End If
End Sub

Private Sub cmdFind_Click()
cmdFindFirst.Enabled = True
Label2.Visible = True
cboField.Visible = True
cmdFindFirst.Visible = True
cmdFindNext.Visible = True
Command2.Visible = True
Command3.Visible = True
picStatBoxReg.Visible = False
End Sub


Private Sub cmdFindFirst_Click()
On Error Resume Next

flag = 0
Adodc1.Recordset.MoveFirst


If a_name <> "" Then
    
    'MsgBox a_name
    
    Do While Not (Adodc1.Recordset.EOF) '& Adodc1.Recordset.Fields(3) = UCase(Trim("a_name")))
        'MsgBox UCase(Adodc1.Recordset.Fields(3)) + "ASHU" + UCase(Trim(a_name))
        If UCase(Adodc1.Recordset.Fields(2)) <> UCase(Trim(a_name)) Then
            'Call proc
            Adodc1.Recordset.MoveNext
        Else
            flag = 1
            Exit Do
        End If
    Loop

ElseIf grn <> "" Then

    Do While Adodc1.Recordset.EOF <> True '& Adodc1.Recordset.Fields(3) = UCase(Trim("a_name")))
        If UCase(Adodc1.Recordset.Fields(3)) <> UCase(Trim(grn)) Then
            'Call proc
            Adodc1.Recordset.MoveNext
        Else
            flag = 2
            Exit Do
        End If
    Loop

ElseIf jv <> "" Then

    Do While Adodc1.Recordset.EOF <> True '& Adodc1.Recordset.Fields(3) = UCase(Trim("a_name")))
        If UCase(Adodc1.Recordset.Fields(4)) <> UCase(Trim(jv)) Then
            'Call proc
            Adodc1.Recordset.MoveNext
        Else
            flag = 3
            Exit Do
        End If
    Loop

ElseIf amt_dr <> "" Then

    Do While Adodc1.Recordset.EOF <> True '& Adodc1.Recordset.Fields(3) = UCase(Trim("a_name")))
        If UCase(Adodc1.Recordset.Fields(6)) <> UCase(Trim(amt_dr)) Then
            'Call proc
            Adodc1.Recordset.MoveNext
        Else
            flag = 4
            Exit Do
        End If
    Loop

ElseIf amt_cr <> "" Then

    Do While Adodc1.Recordset.EOF <> True '& Adodc1.Recordset.Fields(3) = UCase(Trim("a_name")))
        If UCase(Adodc1.Recordset.Fields(7)) <> UCase(Trim(amt_cr)) Then
            'Call proc
            Adodc1.Recordset.MoveNext
        Else
            flag = 5
            Exit Do
        End If
    Loop
    
ElseIf sup <> "" Then

    Do While Adodc1.Recordset.EOF <> True '& Adodc1.Recordset.Fields(3) = UCase(Trim("a_name")))
        If UCase(Adodc1.Recordset.Fields(5)) <> UCase(Trim(sup)) Then
            'Call proc
            Adodc1.Recordset.MoveNext
        Else
            flag = 6
            Exit Do
        End If
    Loop

'ElseIf dat <> "" Then
'
'    Do While Adodc1.Recordset.EOF <> True '& Adodc1.Recordset.Fields(3) = UCase(Trim("a_name")))
'        If UCase(Adodc1.Recordset.Fields(1)) <> UCase(Trim(dat)) Then
'            'Call proc
'            Adodc1.Recordset.MoveNext
'        Else
'            flag = 7
'            Exit Do
'        End If
'    Loop

End If

'MsgBox flag

'If flag = 1 Then
'Label1.Caption = Adodc1.Recordset.Fields(0)
'Text2.Text = Adodc1.Recordset.Fields(1)
'Text3.Text = Adodc1.Recordset.Fields(2)
'Text4.Text = Adodc1.Recordset.Fields(3)
'Text5.Text = Adodc1.Recordset.Fields(4)
'Text6.Text = Adodc1.Recordset.Fields(5)
'Text7.Text = Adodc1.Recordset.Fields(6)
'Text8.Text = Adodc1.Recordset.Fields(7)
'Text9.Text = Adodc1.Recordset.Fields(8)
'Else

If flag = 0 Then
    MsgBox ("No Match Found!!!!")

ElseIf flag = 1 Then
Adodc1.RecordSource = "select * from asset_table where asset like '" & Trim(a_name) & "'"
Adodc1.Refresh
cmdFindFirst.Enabled = False

ElseIf flag = 2 Then
Adodc1.RecordSource = "select * from asset_table where grn like '" & Trim(grn) & "'"
Adodc1.Refresh
cmdFindFirst.Enabled = False

ElseIf flag = 3 Then

Adodc1.RecordSource = "select * from asset_table where JV_IDN_CV like '" & Trim(jv) & "'"
Adodc1.Refresh
cmdFindFirst.Enabled = False

ElseIf flag = 4 Then
Adodc1.Refresh
Adodc1.RecordSource = "select * from asset_table where DEBIT like '" & Trim(amt_dr) & "'"
Adodc1.Refresh
cmdFindFirst.Enabled = False

ElseIf flag = 5 Then
Adodc1.RecordSource = "select * from asset_table where CREDIT like '" & Trim(amt_cr) & "'"
Adodc1.Refresh
cmdFindFirst.Enabled = False

ElseIf flag = 6 Then
Adodc1.RecordSource = "select * from asset_table where SUPPLIER like '" & Trim(sup) & "'"
Adodc1.Refresh
cmdFindFirst.Enabled = False

'ElseIf flag = 7 Then
'Adodc1.RecordSource = "select * from asset_table where DATE like '" & Trim(dat) & "'"
'Adodc1.Refresh
'cmdFindFirst.Enabled = False

Else
    cmdFindFirst.Enabled = False
    'Adodc1.Recordset.MoveNext
End If

End Sub

Private Sub cmdFindNext_Click()
flag = 0
If Not (Adodc1.Recordset.EOF) Then
Adodc1.Recordset.MoveNext
End If
'Adodc1.Recordset.MoveLast
'MsgBox Adodc1.Recordset.EOF
'MsgBox Adodc1.Recordset.Fields(3)
If a_name <> "" Then
   ' If (Adodc1.Recordset.EOF) = True Then
     '   cmdFindNext.Enabled = False
    'End If
    
    'MsgBox a_name
    
    Do While Not (Adodc1.Recordset.EOF) '& Adodc1.Recordset.Fields(3) = UCase(Trim("a_name")))
        'MsgBox UCase(Adodc1.Recordset.Fields(3)) + "ASHU" + UCase(Trim(a_name))
        If UCase(Adodc1.Recordset.Fields(2)) <> UCase(Trim(a_name)) Then
            'Call proc
            Adodc1.Recordset.MoveNext
        Else
            flag = 1
            Exit Do
        End If
    Loop

ElseIf grn <> "" Then
    'If (Adodc1.Recordset.EOF) = True Then
       ' cmdFindNext.Enabled = False
   ' End If

    Do While Adodc1.Recordset.EOF <> True '& Adodc1.Recordset.Fields(3) = UCase(Trim("a_name")))
        If UCase(Adodc1.Recordset.Fields(3)) <> UCase(Trim(grn)) Then
            'Call proc
            Adodc1.Recordset.MoveNext
        Else
            flag = 1
            Exit Do
        End If
    Loop

ElseIf jv <> "" Then
   ' If (Adodc1.Recordset.EOF) = True Then
   '     cmdFindNext.Enabled = False
   ' End If

    Do While Adodc1.Recordset.EOF <> True '& Adodc1.Recordset.Fields(3) = UCase(Trim("a_name")))
        If UCase(Adodc1.Recordset.Fields(4)) <> UCase(Trim(jv)) Then
            'Call proc
            Adodc1.Recordset.MoveNext
        Else
            flag = 1
            Exit Do
        End If
    Loop

ElseIf amt_dr <> "" Then
    'If (Adodc1.Recordset.EOF) = True Then
    '    cmdFindNext.Enabled = False
    'End If

    Do While Adodc1.Recordset.EOF <> True '& Adodc1.Recordset.Fields(3) = UCase(Trim("a_name")))
        If UCase(Adodc1.Recordset.Fields(6)) <> UCase(Trim(amt_dr)) Then
            'Call proc
            Adodc1.Recordset.MoveNext
        Else
            flag = 1
            Exit Do
        End If
    Loop

ElseIf amt_cr <> "" Then
    
    'If (Adodc1.Recordset.EOF) = True Then
    '    cmdFindNext.Enabled = False
    'End If

    Do While Adodc1.Recordset.EOF <> True '& Adodc1.Recordset.Fields(3) = UCase(Trim("a_name")))
        If UCase(Adodc1.Recordset.Fields(7)) <> UCase(Trim(amt_cr)) Then
            'Call proc
            Adodc1.Recordset.MoveNext
        Else
            flag = 1
            Exit Do
        End If
    Loop

End If

If flag = 1 Then
    If Adodc1.Recordset.EOF = True Then
        cmdFindNext.Enabled = False
    End If
    'MsgBox ("Search Complete!!!")
    cmdFindFirst.Enabled = True
Else
    MsgBox ("Search Complete!!!!")
End If

End Sub

Private Sub cmdFirst_Click()
'cmdNext.Enabled = True
'cmdLast.Enabled = True
'cmdPrevious.Enabled = False
'cmdFirst.Enabled = False

If Adodc1.Recordset.BOF = False Then
   Adodc1.Recordset.MoveFirst
End If
CheckNavigation
End Sub

Private Sub cmdLast_Click()
'cmdNext.Enabled = False
'cmdLast.Enabled = False
'cmdPrevious.Enabled = True
'cmdFirst.Enabled = True

If Adodc1.Recordset.EOF = False Then
    Adodc1.Recordset.MoveLast
End If
CheckNavigation
End Sub

Private Sub cmdNext_Click()
'cmdFirst.Enabled = True
'cmdPrevious.Enabled = True
'cmdLast.Enabled = True

If Adodc1.Recordset.EOF = False Then
    'cmdNext.Enabled = True
    Adodc1.Recordset.MoveNext
''Else
    'cmdNext.Enabled = False
    'cmdLast.Enabled = False
    'cmdPrevious.Enabled = True
    'cmdFirst.Enabled = True
End If
CheckNavigation
End Sub
Private Sub CheckNavigation()
  'This will check which navigation button can be
  'accessed when you navigate the recordset through
  'Datagrid control or navigation button itself
  With Adodc1.Recordset
   'If we have at least two record...
   If (.RecordCount > 1) Then
      'BOF = Begin Of Recordset
      If (.BOF) Or _
         (.AbsolutePosition = 1) Then
          cmdFirst.Enabled = False
          cmdPrevious.Enabled = False
          cmdNext.Enabled = True
          cmdLast.Enabled = True
      'EOF = End Of Recordset
      ElseIf (.EOF) Or _
          (.AbsolutePosition = .RecordCount) Then
          cmdNext.Enabled = False
          cmdLast.Enabled = False
          cmdFirst.Enabled = True
          cmdPrevious.Enabled = True
      Else
          cmdFirst.Enabled = True
          cmdPrevious.Enabled = True
          cmdNext.Enabled = True
          cmdLast.Enabled = True
      End If
   Else
      cmdFirst.Enabled = False
      cmdPrevious.Enabled = False
      cmdNext.Enabled = False
      cmdLast.Enabled = False
   End If
 End With
End Sub
Private Sub cmdPrevious_Click()
'cmdNext.Enabled = True
'cmdLast.Enabled = True

If Adodc1.Recordset.BOF = False Then
    'cmdPrevious.Enabled = True
    Adodc1.Recordset.MovePrevious
'Else
    'cmdPrevious.Enabled = False
    'cmdFirst.Enabled = False
    'cmdNext.Enabled = True
End If
CheckNavigation
End Sub

Private Sub Command1_Click()

Adodc1.Recordset.update
update = 1
MsgBox ("Record Added/Updated!!!!!")
'End If
Adodc1.Recordset.MoveFirst

End Sub

Private Sub Command2_Click()
Label2.Visible = False
cboField.Visible = False
cmdFindFirst.Visible = False
cmdFindNext.Visible = False
Command2.Visible = False
Command3.Visible = False
picStatBoxReg.Visible = True
End Sub

Private Sub Command3_Click()
a_name = ""
grn = ""
jv = ""
amt_dr = ""
amt_cr = ""
sup = ""
'dat = ""
flag = 0
cmdFindFirst.Enabled = True
If cboField.ListIndex = 0 Then
    a_name = InputBox("Enter the asset name", "Asset Name", "")
ElseIf cboField.ListIndex = 1 Then
    amt_dr = InputBox("Enter the debit Amount", "Amount (Debit)", "")
ElseIf cboField.ListIndex = 2 Then
    amt_cr = InputBox("Enter the credit Amount", "Amount (Credit)", "")
ElseIf cboField.ListIndex = 3 Then
    grn = InputBox("Enter the GRN", "GRN", "")
ElseIf cboField.ListIndex = 4 Then
    jv = InputBox("Enter the JV_DN_CV", "JV_DN_CV", "")
ElseIf cboField.ListIndex = 5 Then
    sup = InputBox("Enter the Supplier Name", "Supplier", "")
'ElseIf cboField.ListIndex = 6 Then
'    dat = InputBox("Enter the Date", "Date", "")
Else
    MsgBox ("Select The Search Criteria First!!")
End If


End Sub

Private Sub Command4_Click()
Adodc1.RecordSource = "select * from asset_table"
Adodc1.Refresh
End Sub



Private Sub Form_Load()
update = 0
cboField.List(0) = ("Asset Name")
cboField.List(1) = ("Debit Amount")
cboField.List(2) = ("Credit Amount")
cboField.List(3) = ("GRN")
cboField.List(4) = ("JV_DN_CV")
cboField.List(5) = ("Supplier")
'cboField.List(6) = ("Date")
End Sub


Private Sub Timer1_Timer()
If Shape1.BorderColor = vbBlack Then
Shape1.BorderColor = vbRed
Else
Shape1.BorderColor = vbBlack
If Shape1.BorderColor = vbRed Then
Shape1.BorderColor = vbBlue
Else
Shape1.BorderColor = vbBlack
End If
End If
End Sub
