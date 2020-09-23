VERSION 4.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PAK Directory/Extraction"
   ClientHeight    =   3855
   ClientLeft      =   1590
   ClientTop       =   1470
   ClientWidth     =   6690
   Height          =   4260
   Left            =   1530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Top             =   1125
   Width           =   6810
   Begin Threed.SSPanel pnlPercent 
      Height          =   1335
      Left            =   1440
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   2355
      _StockProps     =   15
      BackColor       =   12632256
      BevelInner      =   1
      Begin Threed.SSPanel pnlDisplay 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   661
         _StockProps     =   15
         BackColor       =   12632256
         BevelOuter      =   1
         FloodType       =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   255
         Left            =   75
         TabIndex        =   7
         Top             =   75
         Width           =   3825
         _Version        =   65536
         _ExtentX        =   6747
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "  Saving file ..."
         BackColor       =   12632256
         Alignment       =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Output to : "
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   3735
      End
   End
   Begin VB.ListBox lstEntries 
      Height          =   3375
      Left            =   0
      TabIndex        =   11
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton cmdChooseFile 
      Caption         =   "&File"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblFileSize 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Caption         =   "File size"
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblNumEnt 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblOffset 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Caption         =   "Entries"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Caption         =   "Offset"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   0
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      CancelError     =   -1  'True
      DialogTitle     =   "Choose PAK file"
      Filter          =   "Pak Files|*.pak"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'Option Explicit



Function OpenPAKFile(filename As String)
Dim FFe As Integer, loopvar As Integer, curentry As Integer
Dim ReadString As String * 64
FFe = FreeFile
Open filename For Random As #FFe Len = Len(PakHeader)
Get #FFe, , PakHeader
Close #FFe
If PakHeader.magic <> "PACK" Then
    FFe = MsgBox("Invalid Pack File", vbOKOnly + vbInformation)
    Exit Function
End If
lblOffset.Caption = PakHeader.diroffset
lblNumEnt.Caption = PakHeader.dirsize
ReDim Preserve StoreEntry((PakHeader.dirsize / 64))
FFe = FreeFile
curentry = 1
Open filename For Binary As #FFe
Get #FFe, PakHeader.diroffset + 1, PakEntry
lstEntries.Clear
For loopvar = 1 To ((PakHeader.dirsize / 64) - 1)
    lstEntries.AddItem PakEntry.filename
    StoreEntry(curentry).filename = PakEntry.filename
    StoreEntry(curentry).offset = PakEntry.offset
    StoreEntry(curentry).size = PakEntry.size
    curentry = curentry + 1
    Get #FFe, , PakEntry
Next
Close #FFe
End Function

Function saveentry(filename As String, entryfilename As String, listnumber As Integer)
Dim FFe As Integer
Dim FFd As Integer
Dim WriteData As String
On Error GoTo CFU
Me.MousePointer = 11
lstEntries.Enabled = False
pnlPercent.Visible = True
pnlDisplay.FloodPercent = 0
Label1.Caption = "Output to : " & entryfilename
FFe = FreeFile
DoEvents
Open filename For Binary As #FFe
FFd = FreeFile
totget = StoreEntry(listnumber).size
Open entryfilename For Binary As #FFd
Seek #FFe, StoreEntry(listnumber).offset + 1
While totget > 0
    If totget > 64000 Then
        WriteData = Input$(64000, #FFe)
        totget = totget - 64000
    Else
        WriteData = Input$(totget, #FFe)
        totget = 0
    End If
    Put #FFd, , WriteData
    DoEvents
    loopvar = (StoreEntry(listnumber).size - totget)
    pnlDisplay.FloodPercent = ((loopvar / StoreEntry(listnumber).size) * 100)
Wend
GoTo OfHere
CFU:
X = MsgBox("Error : " & Error & " saving file.", vbOKOnly + vbInformation)
Resume OfHere
OfHere:
Close #FFe
Close #FFd
Me.MousePointer = 0
lstEntries.Enabled = True
pnlPercent.Visible = False
End Function


Private Sub cmdChooseFile_Click()
On Local Error GoTo NoFile
CommonDialog1.Action = 1
OpenPAKFile CommonDialog1.filename
GoTo Here
NoFile:
Resume Here
Here:
On Error GoTo 0
End Sub


Private Sub lstEntries_Click()
lblFileSize.Caption = StoreEntry(lstEntries.ListIndex + 1).size
End Sub


Private Sub lstEntries_DblClick()
Dim thisname As String, curpos As Integer
'Directory Path
basedir = Left$(CommonDialog1.filename, Len(CommonDialog1.filename) - Len(CommonDialog1.FileTitle))
If Right$(basedir, 1) <> "\" Then basedir = basedir & "\"
thisname = lstEntries.Text
curpos = 1
FoundItem = " "
While FoundItem <> ""
    old_curpos = curpos
    curpos = InStr(old_curpos, thisname, "/")
    If curpos > 0 Then
        FoundItem = Mid$(thisname, old_curpos, (curpos - old_curpos))
        basedir = basedir & FoundItem
        findthis = Dir$(basedir, vbDirectory)
        If findthis = "" Then
            On Error GoTo skipdirectory
            MkDir basedir
        End If
        If Right$(basedir, 1) <> "\" Then basedir = basedir & "\"
    Else
        FoundItem = ""
    End If
    curpos = curpos + 1
Wend
GoTo nk
skipdirectory:
Resume nk
nk:
curpos = Len(thisname)
While curpos > 0 And Mid$(thisname, curpos, 1) <> "/"
    curpos = curpos - 1
Wend
thisname = Right$(thisname, Len(thisname) - (curpos))
saveentry CommonDialog1.filename, basedir & thisname, lstEntries.ListIndex + 1
End Sub


Private Sub List1_Click()

End Sub


