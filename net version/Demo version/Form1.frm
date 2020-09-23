VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "An example to demonstrate the spider pak system - made by spider (spider_1027@btopenworld.com)"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   725
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   9720
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Add a file"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5415
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   14552
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   3969
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Open Pak"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1575
      Left            =   1200
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2778
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   2
      Text            =   "Text4"
      Top             =   2040
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   10815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Spider.pak"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Double click a file to extract it"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   6840
      Width           =   10815
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'>>>> Error codes that our dll could return
    'PAK_NOT_FOUND As String
    'FILE_NOT_FOUND As String
    'EMPTY_PAK As String
    'NOT_VALID_PAK As String
    'DONE_OPERATION As String
    'NO_PAK_LOADED As String
    'PATH_FILE_ACCESS_ERROR As String
    'OBJECT_OUT_OF_BOUNDS As String

Public PAK_SYS As New PAK_SYSTEM.PAK_SYSTEM_CLASS


Dim i As Long

Private Sub Command4_Click()
    CD.InitDir = App.Path
    CD.ShowOpen
    If CD.FileName = "" Then Exit Sub

'>>>>
    If CHECK_IF_ERROR(PAK_SYS.OPENPAK(CD.FileName)) = 0 Then
    Call DRAW_LIST_VIEW
    Label2.Caption = "[ " & CD.FileTitle & " ]"
    Else
    Label2.Caption = "[ No Pak Loaded ]"
    ListView1.ListItems.Clear
    End If
End Sub



Private Sub Command7_Click()
    CD.ShowOpen
'>>>>
    CHECK_IF_ERROR (PAK_SYS.ADD_FILE(CD.FileTitle, CD.FileName))
    Call DRAW_LIST_VIEW
End Sub

Private Sub ListView1_DblClick()
'>>>>
    CHECK_IF_ERROR (PAK_SYS.EXTRACT_FILE(ListView1.SelectedItem.Text, App.Path & "\" & ListView1.SelectedItem.Text))
End Sub

Private Sub DRAW_LIST_VIEW()
Dim HYU As Long
'>>>> Make sure return_num_loaded will not return an error to our long
    HYU = CHECK_IF_ERROR(PAK_SYS.RETURN_NUM_LOADED)
    If HYU = 1 Then Exit Sub
'>>>> Add a list of pak items to our listview
    With frmMAIN.ListView1.ListItems
    .Clear
    For i = 0 To PAK_SYS.RETURN_NUM_LOADED
    .Add .Count + 1, "TT" & 1000 * Rnd(4567), PAK_SYS.RETURN_TITLE(i)
    .Item(.Count).ListSubItems.Add 1, "RR" & 1000 * Rnd(4567), PAK_SYS.RETURN_SIZE(i)
    Next i
    End With
End Sub

Private Function CHECK_IF_ERROR(IS_WHAT As String) As Long

    If IS_WHAT = "PAK_NOT_FOUND" Then GoTo 1
    If IS_WHAT = "FILE_NOT_FOUND" Then GoTo 1
    If IS_WHAT = "EMPTY_PAK" Then GoTo 1
    If IS_WHAT = "NOT_VALID_PAK" Then GoTo 1
    If IS_WHAT = "DONE_OPERATION" Then GoTo 1
    If IS_WHAT = "NO_PAK_LOADED" Then GoTo 1
    If IS_WHAT = "PATH_FILE_ACCESS_ERROR" Then GoTo 1
    If IS_WHAT = "OBJECT_OUT_OF_BOUNDS" Then GoTo 1
    If IS_WHAT = "FILE_ALREADY_EXISTS" Then GoTo 1
    CHECK_IF_ERROR = 0
Exit Function
1
Label3.Caption = "A status message was returned from a dll function <" & IS_WHAT & ">"
CHECK_IF_ERROR = 1
If IS_WHAT = "DONE_OPERATION" Then CHECK_IF_ERROR = 0
If IS_WHAT = "EMPTY_PAK" Then CHECK_IF_ERROR = 0
'If IS_WHAT <> "-1" Then CHECK_IF_ERROR = 0
End Function
