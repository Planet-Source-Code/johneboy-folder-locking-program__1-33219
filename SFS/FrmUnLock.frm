VERSION 5.00
Begin VB.Form FrmUnLock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unlock Folder - SFS"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "FrmUnLock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Top             =   5280
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Top             =   5640
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Unlock Folder"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1150
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Lblfolder 
      AutoSize        =   -1  'True
      Caption         =   "Folder Name Goes Here"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   1950
   End
   Begin VB.Label Label4 
      Caption         =   "Raw Folder Path Label4"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3720
      Picture         =   "FrmUnLock.frx":0CCA
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Raw folder path without GUID Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   2970
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Folder to be Unlocked..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter Password to Unlock Folder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
End
Attribute VB_Name = "FrmUnLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function TDecrypt(iString)
    On Error GoTo uhohs
    Q = ""
    zz = Left(iString, 3)
    a = Left(zz, 1)
    b = Mid(zz, 2, 1)
    c = Mid(zz, 3, 1)
    d = Right(iString, 1)
    a = Int(Asc(a)) 'key 1
    b = Int(Asc(b)) 'key 2
    c = Int(Asc(c)) 'key 3
    d = Int(Asc(d)) 'key 4
    txt = Left(iString, Len(iString) - 1)
    txt2 = Mid(txt, 4, Len(txt)) 'encrypted text
    e = 1
    For X = 1 To Len(txt2)
        f = Mid(txt2, X, 1)
        If e = 1 Then Q = Q & Chr(Asc(f) - a)
        If e = 2 Then Q = Q & Chr(Asc(f) - b)
        If e = 3 Then Q = Q & Chr(Asc(f) - c)
        If e = 4 Then Q = Q & Chr(Asc(f) - d)
        e = e + 1
        If e > 4 Then e = 1
    Next X
    TDecrypt = Q
    Exit Function
uhohs:
    TDecrypt = "Error: Invalid text To Decrypt"
    Exit Function
End Function
Function randomnumber(finished)
Randomize
randomnumber = Int((Val(finished) * Rnd) + 1)
End Function

Private Sub Command1_Click()
On Error Resume Next
If Text1 = Text3 Then
Name "" + Label4.Caption + "" As "" + Label3.Caption + ""
Else
MsgBox "Incorrect Password Entered. Folder will not be unlocked.", vbOKOnly + vbCritical, "Password Error"
End If
Unload Me
End Sub

Private Sub Form_Load()
Dim a, b, c, d
 
 Path = Command
Label4.Caption = Path
 
a = Right(Label4, 48)
b = Len(Label4)
c = Len(Label4) - Len(a)
d = Left(Label4, c)
Label3.Caption = d
Lblfolder.Caption = Right$(Label3.Caption, (Len(Label3.Caption) - InStrRev(Label3.Caption, "\", -1, vbTextCompare)))

Text2.Text = GetPassWord
Text3 = TDecrypt(Text2)
Me.Width = 4380
Me.Height = 2040
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Text1_Change()
Text1.Text = Replace(Text1.Text, " ", "")
If Text1 = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Call Command1_Click
End If
End Sub
