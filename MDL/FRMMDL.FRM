VERSION 4.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MDL Stats"
   ClientHeight    =   5730
   ClientLeft      =   2175
   ClientTop       =   750
   ClientWidth     =   6195
   Height          =   6165
   Left            =   2115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Top             =   375
   Width           =   6315
   Begin VB.CommandButton cmdControls 
      Caption         =   "&Cancel"
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   40
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdControls 
      Caption         =   "&OK"
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   39
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   20
      Left            =   4800
      TabIndex        =   38
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   19
      Left            =   4800
      TabIndex        =   37
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   18
      Left            =   4800
      TabIndex        =   36
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   17
      Left            =   4800
      TabIndex        =   35
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   16
      Left            =   4800
      TabIndex        =   34
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   15
      Left            =   1680
      TabIndex        =   33
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   14
      Left            =   1680
      TabIndex        =   32
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   13
      Left            =   1680
      TabIndex        =   31
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   12
      Left            =   1680
      TabIndex        =   30
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   11
      Left            =   1680
      TabIndex        =   29
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   10
      Left            =   3360
      TabIndex        =   28
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   9
      Left            =   2520
      TabIndex        =   27
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   8
      Left            =   1680
      TabIndex        =   26
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   7
      Left            =   3360
      TabIndex        =   25
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   6
      Left            =   2520
      TabIndex        =   24
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   23
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   4
      Left            =   3360
      TabIndex        =   22
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   3
      Left            =   2520
      TabIndex        =   21
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   20
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   1
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2520
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   2010
      Left            =   3000
      Pattern         =   "*.MDL"
      TabIndex        =   2
      Top             =   0
      Width           =   3135
   End
   Begin VB.DirListBox Dir1 
      Height          =   1605
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label lblFileName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   840
      TabIndex        =   42
      Top             =   2040
      Width           =   5295
   End
   Begin VB.Label lblStuff 
      BackStyle       =   0  'Transparent
      Caption         =   "Filename"
      Height          =   255
      Index           =   15
      Left            =   0
      TabIndex        =   41
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblStuff 
      BackStyle       =   0  'Transparent
      Caption         =   "Bounding Radius"
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   17
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblStuff 
      BackStyle       =   0  'Transparent
      Caption         =   "Numskins"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   16
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblStuff 
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      Height          =   255
      Index           =   14
      Left            =   0
      TabIndex        =   15
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblStuff 
      BackStyle       =   0  'Transparent
      Caption         =   "Flags"
      Height          =   255
      Index           =   13
      Left            =   3120
      TabIndex        =   14
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblStuff 
      BackStyle       =   0  'Transparent
      Caption         =   "Sync Type"
      Height          =   255
      Index           =   12
      Left            =   0
      TabIndex        =   13
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label lblStuff 
      BackStyle       =   0  'Transparent
      Caption         =   "NumFrames"
      Height          =   255
      Index           =   11
      Left            =   3120
      TabIndex        =   12
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label lblStuff 
      BackStyle       =   0  'Transparent
      Caption         =   "NumTris"
      Height          =   255
      Index           =   10
      Left            =   0
      TabIndex        =   11
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label lblStuff 
      BackStyle       =   0  'Transparent
      Caption         =   "NumVerts"
      Height          =   255
      Index           =   9
      Left            =   3120
      TabIndex        =   10
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label lblStuff 
      BackStyle       =   0  'Transparent
      Caption         =   "Skin Height"
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblStuff 
      BackStyle       =   0  'Transparent
      Caption         =   "Skin Width"
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblStuff 
      BackStyle       =   0  'Transparent
      Caption         =   "Eye Position [x,y,z]"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   7
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label lblStuff 
      BackStyle       =   0  'Transparent
      Caption         =   "Origin [x,y,z]"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblStuff 
      BackStyle       =   0  'Transparent
      Caption         =   "Scale [x,y,z]"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblStuff 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblStuff 
      BackStyle       =   0  'Transparent
      Caption         =   "Ident"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
Dim ThisFile As Integer
Dim Loaded As Boolean, Dirty As Boolean

Private Sub cmdControls_Click(Index As Integer)
Dim Ffe As Integer
If Index = 0 Then
    ReadRecord.id = txtOther(0).Text
    ReadRecord.Version = txtOther(1).Text
    ReadRecord.[scale].x = Val(txtOther(2).Text)
    ReadRecord.[scale].y = Val(txtOther(3).Text)
    ReadRecord.[scale].z = Val(txtOther(4).Text)
    ReadRecord.[origin].x = Val(txtOther(5).Text)
    ReadRecord.[origin].y = Val(txtOther(6).Text)
    ReadRecord.[origin].z = Val(txtOther(7).Text)
    ReadRecord.[offsets].x = Val(txtOther(8).Text)
    ReadRecord.[offsets].y = Val(txtOther(9).Text)
    ReadRecord.[offsets].z = Val(txtOther(10).Text)
    ReadRecord.skinheight = txtOther(11).Text
    ReadRecord.numtris = txtOther(12).Text
    ReadRecord.synctype = txtOther(13).Text
    ReadRecord.Size.[scale] = txtOther(14).Text
    ReadRecord.numskins = txtOther(15).Text
    ReadRecord.skinwidth = txtOther(16).Text
    ReadRecord.numverts = txtOther(17).Text
    ReadRecord.numframes = txtOther(18).Text
    ReadRecord.flags = txtOther(19).Text
    ReadRecord.radius.[scale] = txtOther(20).Text
    Ffe = FreeFile
    Open File1.Path & "\" & File1.filename For Random As #Ffe Len = Len(ReadRecord)
    Put #Ffe, 1, ReadRecord
    Close #Ffe
    Dirty = False
    cmdControls(0).Enabled = False
    cmdControls(1).Enabled = False
End If
If Index = 1 Then
    Dirty = False
    File1.ListIndex = ThisFile
End If
End Sub

Private Sub Dir1_Change()
If Right$(Dir1.Path, 1) <> "\" Then
    File1.Path = Dir1.Path & "\"
Else
    File1.Path = Dir1.Path
End If
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub


Private Sub File1_Click()
Dim Ffe As Integer
If Loaded Then
    If Dirty Then
        Ffe = MsgBox("Do you wish to save changes to current model?", vbYesNoCancel + vbQuestion)
        If Ffe = vbCancel Then Exit Sub
        If Ffe = vbYes Then
            Call cmdControls_Click(0)
        End If
    End If
End If
Ffe = FreeFile
Open File1.Path & "\" & File1.filename For Random As #Ffe Len = Len(ReadRecord)
Get #Ffe, 1, ReadRecord
Close #Ffe
txtOther(0).Text = ReadRecord.id
txtOther(1).Text = ReadRecord.Version
Call Normalise(ReadRecord.scale, 2)
Call Normalise(ReadRecord.origin, 5)
Call Normalise(ReadRecord.offsets, 8)
txtOther(11).Text = ReadRecord.skinheight
txtOther(12).Text = ReadRecord.numtris
txtOther(13).Text = ReadRecord.synctype
txtOther(14).Text = ReadRecord.Size.scale
txtOther(15).Text = ReadRecord.numskins
txtOther(16).Text = ReadRecord.skinwidth
txtOther(17).Text = ReadRecord.numverts
txtOther(18).Text = ReadRecord.numframes
txtOther(19).Text = ReadRecord.flags
txtOther(20).Text = ReadRecord.radius.scale
ThisFile = File1.ListIndex
lblFileName.Caption = File1.List(ThisFile)
Loaded = True
Dirty = False
For Ffe = 2 To 20: txtOther(Ffe).Enabled = True: Next
cmdControls(0).Enabled = False
cmdControls(1).Enabled = False
End Sub


Private Sub Form_Load()
Dim x As Integer
Loaded = False
Dirty = False
ThisFile = -1
Dir1.Path = "C:\TECHNET\SAMPLES\SOFTLIB\CDLLS\MC\PROGS\"
For x = 2 To 20: txtOther(x).Enabled = False: Next
End Sub




Private Sub txtOther_Change(Index As Integer)
If Loaded Then
    Dirty = True
    cmdControls(0).Enabled = True
    cmdControls(1).Enabled = True
End If
End Sub

Private Sub txtOther_KeyPress(Index As Integer, KeyAscii As Integer)
If Index > 1 Then
    If KeyAscii > 12 And KeyAscii < 127 Then
        If KeyAscii <> Asc("-") And KeyAscii <> Asc(".") Then
            If KeyAscii > Asc("9") Or KeyAscii < Asc("0") Then
                KeyAscii = 0
                Beep
            End If
        Else
            If InStr(txtOther(Index).Text, Chr$(KeyAscii)) > 0 Then
                KeyAscii = 0
                Beep
            End If
        End If
    End If
End If
If Index = 12 Or Index = 14 Or Index = 17 Or Index = 18 Or Index = 19 Then
    If KeyAscii = Asc("-") Or KeyAscii = Asc(".") Then
        KeyAscii = 0
        Beep
    End If
End If
If Index = 13 And (KeyAscii > 12 And KeyAscii < 127) And ((KeyAscii <> Asc("0") And KeyAscii <> Asc("1")) Or Len(Trim$(txtOther(Index).Text)) > 0) Then
    KeyAscii = 0
    Beep
End If
End Sub


