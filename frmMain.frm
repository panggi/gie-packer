VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GIE PACKER FOR EXECUTABLE FILE "
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   795
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   17
      Top             =   8220
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000007&
      Caption         =   "Fungsi:"
      ForeColor       =   &H0000FF00&
      Height          =   1815
      Left            =   3120
      TabIndex        =   8
      Top             =   4080
      Width           =   1935
      Begin VB.CommandButton cmdCompress 
         Caption         =   "Kompres"
         Height          =   525
         Left            =   480
         TabIndex        =   11
         Top             =   1200
         Width           =   855
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   360
         Max             =   9
         Min             =   1
         TabIndex        =   10
         Top             =   720
         Value           =   1
         Width           =   1155
      End
      Begin VB.TextBox txtQuality 
         Height          =   285
         Left            =   750
         MaxLength       =   1
         TabIndex        =   9
         Text            =   "1"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000007&
         Caption         =   "9"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   1680
         TabIndex        =   14
         Top             =   720
         Width           =   135
      End
      Begin VB.Label lblQuality 
         BackColor       =   &H80000007&
         Caption         =   "Kualitas:"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000007&
         Caption         =   "1"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   105
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Tambahkan"
      Height          =   345
      Left            =   1800
      TabIndex        =   3
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear"
      Height          =   345
      Left            =   1800
      TabIndex        =   7
      Top             =   5400
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Caption         =   "Daftar File"
      ForeColor       =   &H0000FF00&
      Height          =   1785
      Left            =   0
      TabIndex        =   5
      Top             =   4080
      Width           =   3045
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Hapus"
         Height          =   345
         Left            =   1800
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         Top             =   330
         Width           =   1545
      End
   End
   Begin VB.ListBox lstFiles 
      Height          =   645
      Left            =   1560
      TabIndex        =   4
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   5115
      Begin VB.TextBox Text2 
         Height          =   345
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   330
         Width           =   1725
      End
      Begin VB.CommandButton Command2 
         Caption         =   "About"
         Height          =   435
         Left            =   4200
         TabIndex        =   18
         Top             =   300
         Width           =   735
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   345
         Left            =   3240
         TabIndex        =   16
         Top             =   330
         Width           =   825
      End
      Begin VB.TextBox txtOutputDir 
         Height          =   345
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   330
         Width           =   1725
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000007&
         Caption         =   "Folder Output:"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   390
         Width           =   1155
      End
   End
   Begin GIEPACKER.ReadOutput ReadOutput1 
      Left            =   3300
      Top             =   8220
      _extentx        =   1720
      _extenty        =   1296
   End
   Begin VB.Image Image1 
      Height          =   2880
      Left            =   -720
      Picture         =   "frmMain.frx":170A2
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




Private Sub cmdBrowse_Click()
Dim strResFolder As String
Text2.Visible = False
txtOutputDir.Visible = True
txtOutputDir.Text = ""
strResFolder = BrowseForFolder(hwnd, "Pilih Folder.")

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
        .DialogTitle = "Pilih File Executable"
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
        MsgBox "Error: " & Err.Description, vbCritical, "Tambah"
    End If
    
End Sub

Private Sub cmdClearAll_Click()
  lstFiles.Clear
  List1.Clear
  Text1.Text = ""
End Sub

Private Sub cmdCompress_Click()
On Error Resume Next
Dim strCommand, strOption As String
Dim strRun, sOutput, TMP As String
Dim i, j, n As Integer
  
  strCommand = ""
  strCommand = "-" & txtQuality.Text & " "
  strOption = ""
  
  
  If Right(txtOutputDir.Text, 1) <> "\" Then
       sOutput = txtOutputDir.Text & "\"
    Else
       sOutput = txtOutputDir.Text
    End If
    
         Me.MousePointer = 11 'busy
     For i = 0 To List1.ListCount - 1
         strRun = "gie " & strCommand & strOption & "-o " & Chr(34) & _
                  sOutput & List1.List(i) & Chr(34) & Chr(32) _
                  & Chr(34) & lstFiles.List(i) & Chr(34)
         
         
         'compressing...
         Text1.Text = "" 'clear
        Next i
        
   Shell (acak.exe - q - i - a & sOutput)
     'Complete!
     Me.MousePointer = 0 'default
  End If
  
Exit Sub
End Sub


Private Sub cmdDelete_Click()
Dim i As Integer

  If List1.ListCount = 0 Then
     MsgBox "Pilih file untuk dihapus.", vbExclamation, "Hapus"
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
FrmAcak.Show
End Sub

Private Sub Command2_Click()
frmAbout.Show
End Sub

Private Sub Command3_Click()

End Sub

Private Sub txtQuality_Change()
  HScroll1.value = txtQuality.Text
End Sub

Private Sub txtQuality_KeyPress(KeyAscii As Integer)
  'if input 1 to 9 or Backspace
If KeyAscii >= 49 And KeyAscii <= 57 Or KeyAscii = 8 Then
   Else
   KeyAscii = 0
End If
End Sub

Private Sub HScroll1_Change()
  txtQuality.Text = HScroll1.value
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
    txtOutputDir.Visible = False
    Text2.Enabled = False
    App.TaskVisible = False
    If App.PrevInstance = True Then End

    Set c = New cFileDialog
    
    txtOutputDir.Text = App.Path
    'copy it to WINDOWS Directory, so can run it anyway!
    If FileExists(Get_WinPath & "gie.exe") = False Then
       DropFile "CUSTOM", 101, Get_WinPath & "gie.exe"

       End If
    If FileExists(Get_WinPath & "acak.exe") = False Then
       DropFile "CUSTOM", 102, Get_WinPath & "acak.exe"
       
       End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
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



