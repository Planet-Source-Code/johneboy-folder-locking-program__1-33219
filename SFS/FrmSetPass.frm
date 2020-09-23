VERSION 5.00
Begin VB.Form FrmSetPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Password - SFS"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   Icon            =   "FrmSetPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3960
      TabIndex        =   4
      Top             =   7200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   4815
      Begin VB.CommandButton Command1 
         Caption         =   "&Set Password"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
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
         TabIndex        =   8
         Top             =   840
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password"
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
         TabIndex        =   7
         Top             =   240
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "FrmSetPass.frx":0CCA
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmSetPass.frx":1994
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   3705
      End
   End
End
Attribute VB_Name = "FrmSetPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function TEncrypt(iString)
On Error GoTo uhoh
    Q = ""
    a = randomnumber(9) + 32
    b = randomnumber(9) + 32
    c = randomnumber(9) + 32
    d = randomnumber(9) + 32
    Q = Chr(a) & Chr(c) & Chr(b)
    e = 1
    For X = 1 To Len(iString)
        f = Mid(iString, X, 1)
        If e = 1 Then Q = Q & Chr(Asc(f) + a)
        If e = 2 Then Q = Q & Chr(Asc(f) + c)
        If e = 3 Then Q = Q & Chr(Asc(f) + b)
        If e = 4 Then Q = Q & Chr(Asc(f) + d)
        e = e + 1
        If e > 4 Then e = 1
    Next X
    Q = Q & Chr(d)
    TEncrypt = Q
    Exit Function
uhoh:
    TEncrypt = "Error: Invalid text To Encrypt"
    Exit Function
End Function
Function randomnumber(finished)
Randomize
randomnumber = Int((Val(finished) * Rnd) + 1)
End Function

Private Sub Command1_Click()
If Text1 = Text3 Then
SetStringValue "HKEY_CURRENT_USER\Software\SFS", "Password", "" + Text2.Text + ""
Unload Me
FrmHowTo.Show
Else
MsgBox "The password does not match. Please retype the password.", vbOKOnly + vbCritical, "Password Error"
End If
End Sub

Private Sub Text1_Change()
Text1.Text = Replace(Text1.Text, " ", "")
Text2 = TEncrypt(Text1)
If Text1 = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub

