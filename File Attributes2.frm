VERSION 5.00
Begin VB.Form FAttrib 
   Caption         =   "Form1"
   ClientHeight    =   8475
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   1035
      Left            =   240
      TabIndex        =   11
      Top             =   5160
      Width           =   6135
   End
   Begin VB.CommandButton Setall 
      Caption         =   "Set File List Attributes"
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton Set1 
      Caption         =   "Set File Attributes"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   7920
      Width           =   2295
   End
   Begin VB.CheckBox Atbs 
      Caption         =   "Archive"
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   8
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CheckBox Atbs 
      Caption         =   "Directory"
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   7
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CheckBox Atbs 
      Caption         =   "System"
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   6
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CheckBox Atbs 
      Caption         =   "Hidden"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CheckBox Atbs 
      Caption         =   "Read Only"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CheckBox Atbs 
      Caption         =   "Normal"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   6480
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Height          =   4770
      Hidden          =   -1  'True
      Left            =   3600
      System          =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   4365
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   2880
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "FAttrib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************
'*                                                       *
'*        Designed by Richard Strom                      *
'*             Free for to use or modify as you like     *
'*         Except please leave this credit as is         *
'*            if you like it,find bugs,ect               *
'*               send email to venusflitrap@netzero.net  *
'*    Warranty:                                          *
'*       as is', without warranties as to performance    *
'*       fitness, merchantability,or any other warranty  *
'*                                                       *
'*********************************************************
Dim Pn As String, Fn As String, Cr As String
Dim T As Integer, I As Integer
Dim Flag As Boolean, DirProp As Boolean
Dim FSys As New FileSystemObject
Private Sub Atbs_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Fn = "" Then
    If DirProp = True Then
      Atbs(4).Value = 1
    Else
      Atbs(4).Value = 0
    End If
    X = MsgBox("Changing Folder Attributes" + Cr + "       is not Available", 32, "Set Attributes")
    Attrib
    Exit Sub
  End If
  If Atbs(0).Value = 1 Then
    ClrCks
    Atbs(0).Value = 1
    Exit Sub
  End If
  Atbs(0).Value = 0
End Sub
Private Sub Dir1_Change()
  Atbs(0).Enabled = False
  Atbs(2).Enabled = False
  Atbs(3).Enabled = False
  Atbs(4).Enabled = True
  File1.Path = Dir1.Path
  Pn = Dir1.Path
  Fn = ""
  Flag = True
  Attrib
End Sub
Private Sub Dir1_Click()
  Dir1.Path = Dir1.List(Dir1.ListIndex)
End Sub
Private Sub Drive1_Change()
  Atbs(0).Enabled = False
  Atbs(2).Enabled = False
  Atbs(3).Enabled = False
  Atbs(4).Enabled = True
  Dir1.Path = Drive1.List(Drive1.ListIndex)
  File1.Path = Dir1.Path
  Pn = Dir1.Path
End Sub
Private Sub File1_Click()
  Atbs(4).Enabled = False
  Atbs(0).Enabled = True
  Atbs(2).Enabled = True
  Atbs(3).Enabled = True
  Pn = Dir1.Path
  Fn = File1.List(File1.ListIndex)
  Attrib
End Sub
Private Sub Form_Load()
    Cr = Chr(13) + Chr(10)
    Flag = False
    Drive1_Change
End Sub
Public Sub Attrib()
  If Right(Pn, 1) <> "\" Then
    Pn = Pn + "\"
  End If
  T = GetAttr(Pn + Fn)
  ClrCks
  List1.Clear
  List1.AddItem Pn + Fn + "   = Attributes #" + Str(T)
  If T = 0 Then
    Atbs(0).Value = 1
  End If
  If T >= 32 Then
    Atbs(5).Value = 1
    T = T - 32
  End If
  If T >= 16 Then
    Atbs(4).Value = 1
    T = T - 16
  End If
  If Atbs(3).Value = 1 Then DirProp = True
  If T >= 2 Then
    Atbs(2).Value = 1
    T = T - 2
  End If
  If T >= 1 Then
    Atbs(1).Value = 1
  End If
  If Atbs(4).Value = 1 Then
    DirProp = True
    If Len(Pn) > 4 Then
      Set Qn = FSys.GetFolder(Pn)
      Label1.Caption = Qn.DateCreated
      List1.AddItem "Date Created " + Label1.Caption
      Label1.Caption = Qn.DateLastModified
      List1.AddItem "Last Modified " + Label1.Caption
      Label1.Caption = Qn.DateLastAccessed
      List1.AddItem "Last Accessed " + Label1.Caption
      Label1.Caption = Qn.Size
      List1.AddItem "Folder Size in Bytes " + Label1.Caption
    End If
  Else
    DirProp = False
    Set Qn = FSys.GetFile(Pn + Fn)
    Label1.Caption = Qn.DateCreated
    List1.AddItem "Date Created " + Label1.Caption
    Label1.Caption = Qn.DateLastModified
    List1.AddItem "Last Modified " + Label1.Caption
    Label1.Caption = Qn.DateLastAccessed
    List1.AddItem "Last Accessed " + Label1.Caption
    Label1.Caption = Qn.Size
    List1.AddItem "File Size in Bytes " + Label1.Caption
  End If
  Flag = False
End Sub
Public Sub ClrCks()
  For I = 0 To 5
    Atbs(I).Value = 0
  Next I
End Sub
Private Sub Set1_Click()
  If Fn = "" Then
    X = MsgBox("Changing Folder Attributes" + Cr + "       is not Available", 32, "Set Attributes")
    Exit Sub
  End If
  X = MsgBox("Set Attributes of " + Cr + Pn + Fn, 33, "Set Attributes")
  If X <> 1 Then Attrib: Exit Sub
  DoAtr
  Attrib
End Sub
Public Sub DoAtr()
  T = 0
  If Atbs(5).Value = 1 And Atbs(5).Enabled = True Then
    T = T + 32
  End If
  If Atbs(4).Value = 1 And Atbs(4).Enabled = True Then
    T = T + 16
  End If
  If Atbs(3).Value = 1 And Atbs(3).Enabled = True Then
    T = T + 4
  End If
  If Atbs(2).Value = 1 And Atbs(2).Enabled = True Then
    T = T + 2
  End If
  If Atbs(1).Value = 1 And Atbs(1).Enabled = True Then
    T = T + 1
  End If
  If Atbs(0).Value = 1 And Atbs(0).Enabled = True Then T = 0
 SetAttr Pn + Fn, T
End Sub
Private Sub Setall_Click()
  X = MsgBox("Caution This Will Set Attributes of all Files " + Cr + "in Folder " + Pn, 49, "Set Attributes")
  If X <> 1 Then Attrib: Exit Sub
  For I = 0 To File1.ListCount - 1
    Fn = File1.List(I)
    DoAtr
  Next I
End Sub
