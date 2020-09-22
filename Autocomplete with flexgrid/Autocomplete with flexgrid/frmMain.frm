VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Cd Finder"
   ClientHeight    =   5415
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6495
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin MSFlexGridLib.MSFlexGrid FlexResultado 
         Height          =   3495
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   "Artist                           |          CD Name                                    | Price   | Reference"
      End
      Begin VB.TextBox TextoaBuscar 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox cboBusqueda 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Text To Search:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Search by:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuVote 
      Caption         =   "Please Vote for ME !!"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DB As Database
Public Rs As Recordset

Private Sub cboBusqueda_Click()
    TextoaBuscar.Text = ""
End Sub

Private Sub Form_Load()
With cboBusqueda
    .AddItem "Artist"
    .AddItem "CD Name"
    .ListIndex = 0
End With
If Right(App.Path, 1) = "\" Then
    Set DB = OpenDatabase(App.Path + "DB.mdb")
Else
    Set DB = OpenDatabase(App.Path + "\DB.mdb")
End If
Set Rs = DB.OpenRecordset("ListaCDs")
End Sub
Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuVote_Click()
    MsgBox "Please Vote for ME !!", vbExclamation, "Vote - Daniel Vargas"
End Sub

Private Sub TextoaBuscar_Change()
On Error Resume Next
If cboBusqueda.Text <> "" Then
    FlexResultado.Rows = 1
If FlexResultado.Rows = 1 Then
    FlexResultado.TextMatrix(0, 0) = "Artist"
    FlexResultado.TextMatrix(0, 1) = "CD Name"
    FlexResultado.TextMatrix(0, 2) = "Price"
    FlexResultado.TextMatrix(0, 3) = "Reference"
End If
Select Case cboBusqueda.Text
    Case "Artist"
        AutoComplete TextoaBuscar, FlexResultado, DB, "ListaCDs", "Artista"
    Case "CD Name"
        AutoComplete TextoaBuscar, FlexResultado, DB, "ListaCDs", "Disco"
    Case Else
End Select
Else
    TextoaBuscar = ""
    cboBusqueda.ListIndex = 0
End If
End Sub

