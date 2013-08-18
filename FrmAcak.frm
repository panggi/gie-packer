VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmAcak 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acak File Terkompresi"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "FrmAcak.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Buat Backup"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1500
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Acak File"
      Height          =   495
      Left            =   1860
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buka File"
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   840
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   300
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "File Terkompresi :"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1875
   End
End
Attribute VB_Name = "FrmAcak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Private Sub Command1_Click()
CommonDialog1.Filter = "EXE Files|*.exe"
CommonDialog1.ShowOpen
If Len(Dir$(CommonDialog1.Filename)) > 0 Then
Text1.Text = CommonDialog1.Filename
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim FileData As String
Dim a As Integer

If Len(Dir$(Text1.Text)) = 0 Then
MsgBox "Filenya ga ada!"
Exit Sub
End If

'Backup the file
If Check1.value = 1 Then
FileCopy Text1.Text, Text1.Text & ".gie"
End If

Open Text1.Text For Binary As #1
FileData = Space$(LOF(1))
Get #1, , FileData

'Ganti "UPX0" ama ga ada
If InStr(1, FileData, hex2ascii("55505830")) > 0 Then
Put #1, InStr(1, FileData, hex2ascii("55505830")), hex2ascii("4749452E")
a = a + 4
End If

'Ganti "UPX1" ama ga ada
If InStr(1, FileData, hex2ascii("55505831")) > 0 Then
Put #1, InStr(1, FileData, hex2ascii("55505831")), hex2ascii("00000000")
a = a + 4
End If

'Ganti "UPX!" ama ga ada
If InStr(1, FileData, hex2ascii("55505821")) > 0 Then
Put #1, InStr(1, FileData, hex2ascii("55505821")), hex2ascii("00000000")
a = a + 4
End If

'Ganti ".1.25." ama ..GIE.
If InStr(1, FileData, hex2ascii("00312E323500")) > 0 Then
Put #1, InStr(1, FileData, hex2ascii("00312E323500")), hex2ascii("00004749452E")
a = a + 6
End If

'Ganti ".1.20." ama ..GIE.
If InStr(1, FileData, hex2ascii("00312E323000")) > 0 Then
Put #1, InStr(1, FileData, hex2ascii("00312E323000")), hex2ascii("00004749452E")
a = a + 6
End If

'Ganti "pk...`" ama ok...`
If InStr(1, FileData, hex2ascii("706B00000060")) > 0 Then
Put #1, InStr(1, FileData, hex2ascii("706B00000060")), hex2ascii("6F6B00000060")
a = a + 6
End If

'Ganti "€..à" ama @..À
If InStr(1, FileData, hex2ascii("800000E0")) > 0 Then
Put #1, InStr(1, FileData, hex2ascii("800000E0")), hex2ascii("400000C0")
a = a + 4
End If

'Ganti "@..à." ama @..À.
If InStr(1, FileData, hex2ascii("400000E02E")) > 0 Then
Put #1, InStr(1, FileData, hex2ascii("400000E02E")), hex2ascii("400000C02E")
a = a + 5
End If

'Ganti "i..aé" ama i..`é
If InStr(1, FileData, hex2ascii("69010061E9")) > 0 Then
Put #1, InStr(1, FileData, hex2ascii("69010061E9")), hex2ascii("69010060E9")
a = a + 5
End If

'Ganti "i..aé" ama i..`é
If InStr(1, FileData, hex2ascii("69000061E9")) > 0 Then
Put #1, InStr(1, FileData, hex2ascii("69000061E9")), hex2ascii("69000060E9")
a = a + 5
End If

'Ganti ".Š" ama .ë.ëêëèŠ
If InStr(1, FileData, hex2ascii("109090909090908A")) > 0 Then
Put #1, InStr(1, FileData, hex2ascii("109090909090908A")), hex2ascii("10EB00EBEAEBE88A")
a = a + 9
End If

'Ganti "...`¾.`" ama ..a¾.`
If InStr(1, FileData, hex2ascii("00000060BE0060")) > 0 Then
Put #1, InStr(1, FileData, hex2ascii("00000060BE0060")), hex2ascii("00009061BE0060")
a = a + 7
End If

'Ganti "..aé" ama ..`é
If InStr(1, FileData, hex2ascii("000061E9")) > 0 Then
Put #1, InStr(1, FileData, hex2ascii("000061E9")), hex2ascii("000060E9")
a = a + 4
End If

If a > 1 Then
MsgBox a & " pattern berhasil diacak!!", vbInformation, "Acak File"
Else
MsgBox "Gagal Euy!", vbCritical, ""
End If

Close #1

End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Close #1
End Sub

