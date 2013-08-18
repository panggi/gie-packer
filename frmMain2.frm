VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GIE PROTECTOR v0.2 RELOADED "
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5145
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HasDC           =   0   'False
   Icon            =   "frmMain2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   412
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   343
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1440
      Top             =   6600
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000007&
      Caption         =   "Hasil:"
      ForeColor       =   &H0000FF00&
      Height          =   1485
      Left            =   0
      TabIndex        =   12
      Top             =   7080
      Width           =   5085
      Begin MSComctlLib.ListView ListView1 
         Height          =   1095
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   1931
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nama File"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ukuran"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Dikompres"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Rasio"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   2646
         EndProperty
      End
   End
   Begin VB.TextBox Text1 
      Height          =   795
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   8220
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Options"
      ForeColor       =   &H00C0C0C0&
      Height          =   2265
      Left            =   2040
      TabIndex        =   7
      Top             =   3840
      Width           =   3015
      Begin VB.CheckBox Chkexport 
         BackColor       =   &H00000000&
         Caption         =   "Don't Compress Export section"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "INFO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   20
         Top             =   1800
         Width           =   975
      End
      Begin VB.CheckBox Chkultra 
         BackColor       =   &H00000000&
         Caption         =   "Extreme Compression (very slow)"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox Chkbrute 
         BackColor       =   &H00000000&
         Caption         =   "Try All Methods and Filters (slow)"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox Chkstriprel 
         BackColor       =   &H00000000&
         Caption         =   "Don't Strip Relocations"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox Chkcompres 
         BackColor       =   &H00000000&
         Caption         =   "Don't Compress Resources"
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   2535
      End
      Begin VB.CheckBox ChkLZMA 
         BackColor       =   &H00000000&
         Caption         =   "Use LZMA Algorithm (fast)"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CommandButton cmdCompress 
         BackColor       =   &H80000007&
         Caption         =   "PROTECT"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   480
         MaskColor       =   &H00404040&
         TabIndex        =   8
         Top             =   1800
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "ADD"
      Height          =   345
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "List of File (s)"
      ForeColor       =   &H00C0C0C0&
      Height          =   2265
      Left            =   40
      TabIndex        =   5
      Top             =   3840
      Width           =   1845
      Begin VB.CommandButton cmdDelete 
         Caption         =   "DEL"
         Height          =   345
         Left            =   960
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
         Height          =   1425
         Left            =   120
         MultiSelect     =   2  'Extended
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         Top             =   720
         Width           =   1545
      End
   End
   Begin VB.ListBox lstFiles 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   8040
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   915
      Left            =   40
      TabIndex        =   0
      Top             =   2880
      Width           =   5055
      Begin VB.CommandButton cmdBrowse 
         BackColor       =   &H00000000&
         Caption         =   "......"
         Height          =   345
         Left            =   4320
         TabIndex        =   10
         Top             =   330
         Width           =   585
      End
      Begin VB.TextBox txtOutputDir 
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   330
         Width           =   2805
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000007&
         Caption         =   "Output Folder:"
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   390
         Width           =   1155
      End
   End
   Begin GIEPROTECTOR.ReadOutput ReadOutput1 
      Left            =   3300
      Top             =   8220
      _ExtentX        =   1720
      _ExtentY        =   1296
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "E X I T"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   240
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   2880
      Left            =   0
      Picture         =   "frmMain2.frx":5BBA
      Top             =   0
      Width           =   5865
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private c As cFileDialog
Dim strOutput As String
Dim Temp(6) As String 'Output data(status)
Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
    "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize _
    As Long) As Long
Private SF As String * 255
Private Function SpecialFolder(value)
On Error Resume Next
Dim FolderValue As String
If value = 1 Then
FolderValue = Left(SF, GetSystemDirectory(SF, 255))
End If
If Right(FolderValue, 1) = "\" Then
FolderValue = Left(FolderValue, Len(FolderValue) - 1)
End If
SpecialFolder = FolderValue
End Function

Private Sub cmdBrowse_Click()
Dim strResFolder As String
txtOutputDir.Visible = True
txtOutputDir.Text = ""
strResFolder = BrowseForFolder(hwnd, "Select Output Folder")

If strResFolder <> "" Then
   txtOutputDir.Text = strResFolder
End If
End Sub


Private Sub cmdAdd_Click()
On Error GoTo cmdClassError
Dim sFiles() As String
Dim filecount As Long
Dim sDir As String
Dim i As Long
    
    With c
        .DialogTitle = "Chose Executable File"
        .CancelError = False
        .Filename = "" 'clear
        .hwnd = Me.hwnd
        .flags = OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_ALLOWMULTISELECT
        .InitDir = App.Path
        .Filter = "Executable Files (*.exe)|*.exe"
        .FilterIndex = 1
        .ShowOpen

        If .Filename = "" Then Exit Sub
        .ParseMultiFileName sDir, sFiles(), filecount
        If UBound(sFiles) = 0 Then
           List1.AddItem sFiles(0)
           lstFiles.AddItem .Filename
        Else
           For i = 0 To filecount - 1
               If Mid(sDir, Len(sDir), 1) <> "\" Then
                  lstFiles.AddItem sDir & "\" & sFiles(i)
               Else
                  lstFiles.AddItem sDir & sFiles(i)
               End If
               List1.AddItem sFiles(i)
           Next i
        End If
    End With
    
Exit Sub

cmdClassError:
    If (Err.Number <> 20001) Then
        MsgBox "Error: " & Err.Description, vbCritical, "Add File"
    End If
    
End Sub


Private Sub cmdCompress_Click()
On Error Resume Next
Dim strRun, sOutput, TMP As String
Dim LZMA, Brute, Ultra, Res, Strip, Export As String
Dim i, j, n As Integer
'del previous dependencies
If FileExists(Get_WinPath & "gie.exe") = True Then
      Kill Get_WinPath & "gie.exe"
End If
If FileExists(Get_WinPath & "acak.exe") = True Then
      Kill Get_WinPath & "acak.exe"
End If

'==================================

  
  If ChkLZMA.value = 1 Then
    LZMA = "--lzma "
  Else
    LZMA = "--best "
  End If
  
  If Chkbrute.value = 1 Then
    Brute = "--brute "
  Else
    Brute = ""
  End If
  
  If Chkultra.value = 1 Then
    Ultra = "--ultra-brute "
  Else
    Ultra = ""
  End If
  
  If Chkcompres.value = 1 Then
    Res = "--compress-resources=0 "
  Else
    Res = ""
  End If
  
  If Chkstriprel.value = 1 Then
    Strip = "--strip-relocs=0 "
  Else
    Strip = ""
  End If
  
  If Chkexport.value = 1 Then
    Export = "--compress-exports=0 "
  Else
    Export = ""
  End If
  
  
  
  If Right(txtOutputDir.Text, 1) <> "\" Then
       sOutput = txtOutputDir.Text & "\"
    Else
       sOutput = txtOutputDir.Text
    End If
    
  'if \pack.dll not found
  If FileExists(SpecialFolder(1) & "\pack.dll") = False Then
     MsgBox "Please Restart The Application!", vbExclamation, "Fatal Error"
  Else
  'if \pack.dll found
     ListView1.ListItems.Clear
     ListView1.ColumnHeaders(3).Text = "Packed"
     Me.MousePointer = 11 'busy
     For i = 0 To List1.ListCount - 1
         strRun = SpecialFolder(1) & "\pack.dll -f " & LZMA & Brute & Ultra & Res & Strip & Export & "-o " & Chr(34) & _
                  sOutput & List1.List(i) & Chr(34) & Chr(32) _
                  & Chr(34) & lstFiles.List(i) & Chr(34)
         
        'compressing...
         Text1.Text = "" 'clear
         ReadOutput1.SetCommand = strRun
         ReadOutput1.ProcessCommand
         DoEvents
         j = InStrRev(Text1, "gie:")
         If j = 0 Then 'if j=0 mean not found error.
            TMP = Mid(Text1, 327, Len(Text1) - 327)
            n = StringTokenizer(Trim(TMP) & " ")
            For n = 1 To (n / 6)
                With ListView1.ListItems
                .Add(n).Text = Mid(Temp(6 * n), 1, Len(Temp(6 * n)) - 6) 'Filename
                .item(n).SubItems(1) = FileByteFormat(CLng(Temp(6 * (n - 1) + 1))) 'Size
                .item(n).SubItems(2) = FileByteFormat(CLng(Temp(3 * (n + n - 1)))) 'Packed
                .item(n).SubItems(3) = Temp(3 * (n + n - 1) + 1) 'Ratio
                .item(n).SubItems(4) = "Berez!" 'Status
                End With
            Next n
         Else
            TMP = Mid(Text1, 327, Len(Text1) - 327)
            j = InStrRev(TMP, "AlreadyPacked")
            If j <> 0 Then
               With ListView1.ListItems
                .Add(1).Text = List1.List(i) 'Filename
                .item(1).SubItems(1) = FileByteFormat(FileLen(lstFiles.List(i))) 'Size
                .item(1).SubItems(2) = "0" 'Packed
                .item(1).SubItems(3) = "0" 'Ratio
                .item(1).SubItems(4) = "Already Packed" 'Status
                End With
            End If
            j = InStrRev(TMP, "File exists")
            If j <> 0 Then
               With ListView1.ListItems
                .Add(1).Text = List1.List(i) 'Filename
                .item(1).SubItems(1) = FileByteFormat(FileLen(lstFiles.List(i))) 'Size
                .item(1).SubItems(2) = "0" 'Packed
                .item(1).SubItems(3) = "0" 'Ratio
                .item(1).SubItems(4) = "File exists" 'Status
                End With
            End If
            j = InStrRev(TMP, "Permission denied")
            If j <> 0 Then
               With ListView1.ListItems
                .Add(1).Text = List1.List(i) 'Filename
                .item(1).SubItems(1) = FileByteFormat(FileLen(lstFiles.List(i))) 'Size
                .item(1).SubItems(2) = "0" 'Packed
                .item(1).SubItems(3) = "0" 'Ratio
                .item(1).SubItems(4) = "Permission denied" 'Status
                End With
            End If
         End If
    Shell (SpecialFolder(1) & "\protect.dll -q -i -a " & Chr(34) & sOutput & List1.List(i) & Chr(34))
    Sleep 2, True

    fakesign (sOutput & List1.List(i))

     Next i
     'Complete!
          Me.MousePointer = 0 'default
  End If
  Exit Sub
End Sub
  


Private Sub cmdDelete_Click()
Dim i As Integer
Timer1.Enabled = True
  If List1.ListCount = 0 Then
     MsgBox "Please Chose File To Be Deleted.", vbExclamation, "Delete"
  End If
  
  Do While i < List1.ListCount
      If List1.Selected(i) = True Then
         List1.RemoveItem i
         lstFiles.RemoveItem i
      Else
         i = i + 1
      End If
      DoEvents
  Loop
  
End Sub

Private Sub Command1_Click()
frmAbout.Show
End Sub

Private Sub Label1_Click()
Unload Me
End Sub


Private Sub ReadOutput1_GotChunk(ByVal sChunk As String, ByVal LastChunk As Boolean)
 
 Text1 = Text1 & sChunk
 
End Sub

Private Sub ReadOutput1_Error(ByVal Error As String, LastDLLError As Long)
    MsgBox "Error!" & vbNewLine & _
            "Description: " & Error & vbNewLine & _
            "LastDLLError: " & LastDLLError, vbCritical, "Error"
End Sub

Private Sub Form_Load()
On Error Resume Next
    App.TaskVisible = False
    If App.PrevInstance = True Then End

    Set c = New cFileDialog
    Shell "attrib -s -h -r " & Get_WinPath & "gie.exe"
    Shell "attrib -s -h -r " & Get_WinPath & "acak.exe"
    
    'Copy ke Direktori WINDOWS biar bisa langsung di RUN
    If FileExists(SpecialFolder(1) & "\pack.dll") = False Then
       DropFile "CUSTOM", 101, SpecialFolder(1) & "\pack.dll"
       End If
    If FileExists(SpecialFolder(1) & "\protect.dll") = False Then
       DropFile "CUSTOM", 102, SpecialFolder(1) & "\protect.dll"
       End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set c = Nothing
End Sub

Function StringTokenizer(DATA As String) As Integer
On Error Resume Next
Dim i As Integer, j As Integer, t As Integer
Dim S As String

j = 1 'set data count to zero
t = 1 'set start to zero

For i = 1 To Len(DATA)
    S = Mid$(DATA, i, 1)
    If S = " " Then
       If Trim(Mid$(DATA, t, i - t)) <> "" Then
          Temp(j) = Trim(Mid$(DATA, t, i - t))
          t = i + 1
          j = j + 1
       End If
    End If
Next i

StringTokenizer = j - 1
End Function

Public Function FileByteFormat(FileBytes As Long) As String
On Error Resume Next
Dim nFileNum As Integer
Dim TempNum As Single

If FileBytes > 0 Then
    ' Get file's length
    FileByteFormat = FileBytes / 1024
    
    ' Round number
    TempNum = FileByteFormat - Int(FileByteFormat)
    
    ' Use different scale according to the size of the file
    Select Case Val(FileByteFormat)
        Case Is > 1024 ' Use Mega Byte
            FileByteFormat = Format(FileByteFormat / 1000, "#.##MB")
        Case Else  ' Use Kilo Byte
            ' All values are to round up
            FileByteFormat = Format(FileByteFormat + (1 - TempNum), "###KB")
    End Select
Else
    FileByteFormat = "0KB"
End If

End Function

Private Function DropFile(ResType As String, ResID As Long, _
    OutputPath As String)
On Error Resume Next
Dim DROP() As Byte
DROP = LoadResData(ResID, ResType)
Open OutputPath For Binary As #1
Put #1, , DROP
Close #1
End Function

Private Sub Timer1_Timer()
On Error Resume Next
If txtOutputDir.Text = "" Then
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
Else
    cmdAdd.Enabled = True
    cmdDelete.Enabled = True
End If

If List1.ListCount <> 0 And cmdAdd.Enabled = True Then
    cmdCompress.Enabled = True
Else
    cmdCompress.Enabled = False
End If


End Sub
Public Sub Sleep(Seconds As Single, EventEnable As Boolean)
    On Error GoTo ErrHndl
    Dim OldTimer As Single
    
    OldTimer = Timer
    Do While (Timer - OldTimer) < Seconds
        If EventEnable Then DoEvents
    Loop

    Exit Sub
ErrHndl:
    Err.Clear
End Sub

Public Function hex2ascii(ByVal hextext As String) As String
Dim y As Integer
Dim num As String
Dim value As String
For y = 1 To Len(hextext)
num = Mid(hextext, y, 2)
value = value & Chr(Val("&h" & num))
y = y + 1
Next y

hex2ascii = value
End Function

Public Function fakesign(namafile As String)
On Error Resume Next
Dim FileData As String
Dim ab As Integer
Dim a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p As String
Dim acak As String
Dim RandomName As String
Dim Panjang As Long
Dim ii, aa As Long
Dim alp As String

' Random....
alp = "abcdefghijklmno"
RandomName = ""
Panjang = 1
For ii = 1 To Panjang
Randomize
aa = Int(Rnd * 10)
RandomName = RandomName & Mid(alp, aa, 1)
Next ii

acak = RandomName

Open namafile For Binary As #1
FileData = Space$(LOF(1))
Get #1, , FileData

If acak = "a" Then
Put #1, InStr(1, FileData, hex2ascii("4C6F61644C696272617279410000000000000000000000000000000000000000000000000000000000")), _
hex2ascii("4C6F61644C69627261727941008CCBBA2E2E03DAFC33F633FF4B8EDB8D2E2E2E8EC0B92E2EF3A54A75")
ab = 2
End If

If acak = "b" Then
Put #1, InStr(1, FileData, hex2ascii("4C6F61644C696272617279410000000000000000000000000000000000000000000000000000000000")), _
hex2ascii("4C6F61644C6962726172794100832E2EE22E2EE22EFF00000000000000000000000000000000000000")
ab = 2
End If

If acak = "c" Then
Put #1, InStr(1, FileData, hex2ascii("4C6F61644C696272617279410000000000000000000000000000000000000000000000000000000000")), _
hex2ascii("4C6F61644C69627261727941005B43837B74000F8408000000894314E9000000000000000000000000")
ab = 2
End If

If acak = "d" Then
Put #1, InStr(1, FileData, hex2ascii("4C6F61644C696272617279410000000000000000000000000000000000000000000000000000000000")), _
hex2ascii("4C6F61644C696272617279410068012E4000E801000000C3C300000000000000000000000000000000")
ab = 2
End If

If acak = "e" Then
Put #1, InStr(1, FileData, hex2ascii("4C6F61644C696272617279410000000000000000000000000000000000000000000000000000000000")), _
hex2ascii("4C6F61644C696272617279410068FF6424F06858585858FFD4508B40F205B095F6950F850181BBFF68")
ab = 2
End If

If acak = "f" Then
Put #1, InStr(1, FileData, hex2ascii("4C6F61644C696272617279410000000000000000000000000000000000000000000000000000000000")), _
hex2ascii("4C6F61644C69627261727941005589E583EC0883C4F46A2EA12E2E2E00FFD0E82EFFFFFF0000000000")
ab = 2
End If

If acak = "g" Then
Put #1, InStr(1, FileData, hex2ascii("4C6F61644C696272617279410000000000000000000000000000000000000000000000000000000000")), _
hex2ascii("4C6F61644C69627261727941008BCB85C9742E803A017408ACAE750A4249EBEF47464249EBE9000000")
ab = 2
End If

If acak = "h" Then
Put #1, InStr(1, FileData, hex2ascii("4C6F61644C696272617279410000000000000000000000000000000000000000000000000000000000")), _
hex2ascii("4C6F61644C69627261727941000000000063796200656C6963656E34302E646C6C0000000000000000")
ab = 2
End If

If acak = "i" Then
Put #1, InStr(1, FileData, hex2ascii("4C6F61644C696272617279410000000000000000000000000000000000000000000000000000000000")), _
hex2ascii("4C6F61644C69627261727941004550453A20456E637279707450452056322E323030362E312E313500")
ab = 2
End If

If acak = "j" Then
Put #1, InStr(1, FileData, hex2ascii("4C6F61644C696272617279410000000000000000000000000000000000000000000000000000000000")), _
hex2ascii("4C6F61644C696272617279410083C6148B55FCE92EFFFFFF0000000000000000000000000000000000")
ab = 2
End If

If acak = "k" Then
Put #1, InStr(1, FileData, hex2ascii("4C6F61644C696272617279410000000000000000000000000000000000000000000000000000000000")), _
hex2ascii("4C6F61644C6962726172794100EB2E457865537465616C746820563220536861726577617265200000")
ab = 2
End If

If acak = "l" Then
Put #1, InStr(1, FileData, hex2ascii("4C6F61644C696272617279410000000000000000000000000000000000000000000000000000000000")), _
hex2ascii("4C6F61644C6962726172794100457850722D762E312E322E0000000000000000000000000000000000")
ab = 2
End If

If acak = "m" Then
Put #1, InStr(1, FileData, hex2ascii("4C6F61644C696272617279410000000000000000000000000000000000000000000000000000000000")), _
hex2ascii("4C6F61644C6962726172794100BA2E2E2E2EFFE264114000FF3584114000E840000000000000000000")
ab = 2
End If

If acak = "n" Then
Put #1, InStr(1, FileData, hex2ascii("4C6F61644C696272617279410000000000000000000000000000000000000000000000000000000000")), _
hex2ascii("4C6F61644C69627261727941005E83C664AD50AD5083EE6CAD50AD50AD50AD50AD50E8E70700000000")
ab = 2
End If

If acak = "o" Then
Put #1, InStr(1, FileData, hex2ascii("4C6F61644C696272617279410000000000000000000000000000000000000000000000000000000000")), _
hex2ascii("4C6F61644C696272617279410036812C242E2E2E00C360000000000000000000000000000000000000")
ab = 2
End If

If acak = "p" Then
Put #1, InStr(1, FileData, hex2ascii("4C6F61644C696272617279410000000000000000000000000000000000000000000000000000000000")), _
hex2ascii("4C6F61644C6962726172794100E92E2E2EFF0C2E000000000000000000000000000000000000000000")
ab = 2
End If

If ab > 1 Then
MsgBox "File(s) Protected!!", vbInformation, "Successful"
Else
MsgBox "Something Wrong !!", vbCritical, "Fatal Error"
End If
Close #1
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Close #1
End Sub
