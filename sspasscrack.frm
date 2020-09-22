VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WOH Screen Saver Pass Crack"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   ForeColor       =   &H00FFFFFF&
   Icon            =   "sspasscrack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   1800
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Test Pass"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Pass"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get Pass"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O     P      T"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3720
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.Menu opt1 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu clr1 
         Caption         =   "Clear Text"
      End
      Begin VB.Menu cpy1 
         Caption         =   "Copy Text"
      End
      Begin VB.Menu sv1 
         Caption         =   "Save Text"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clr1_Click()
Text1.Text = ""
End Sub

Private Sub Command1_Click()

Screensavepwd

Text1.Text = "Reading Registry..." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "Encrypted Pass: " & Text2.Text & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "Decrypted Pass: " & Screensavepwd

End Sub

Private Sub Command2_Click()
    PwdChangePassword "SCRSAVE", Me.hwnd, 0, 0
    Me.Show
End Sub

Private Sub Command3_Click()
    Dim bRes As Boolean
    bRes = VerifyScreenSavePwd(Me.hwnd)
    MsgBox bRes
    Me.Show
End Sub

Function Screensavepwd() As String


    'Dim's for the Registry
    Dim lngType As Long, varRetString As Variant
    Dim lngI As Long, intChar As Integer
    'Dim's for the Password decoding
    Dim Ciphertext As String, Key As String
    Dim temp1 As String, temp2 As String
    'Registry Path to the encrypted Password
    varRetString = sdaGetRegEntry("HKEY_CURRENT_USER", _
    "Control Panel\desktop", "ScreenSave_Data", "1")
    
    'the Encrypted Password
    Ciphertext = varRetString
    If Len(Ciphertext) <> 1 Then
        Ciphertext = Left$(varRetString, Len(Ciphertext) - 1)
        Text2.Text = Ciphertext
        'Micro$oft's "Secret" Key
        Key = "48EE761D6769A11B7A8C47F85495975F78D9DA6C59D76B35C57785182A0E52FF00E31B718D3463EB91C3240FB7C2F8E3B6544C3554E7C94928A385110B2C68FBEE7DF66CE39C2DE472C3BB851A123C32E36B4F4DF4A924C8FA78AD23A1E46D9A04CE2BC5B6C5EF935CA8852B413772FA574541A1204F80B3D52302643F6CF10F"
        'XOR every Ciphertextbyte with the Keybyte to get
        'the plaintext
        For i = 1 To Len(Ciphertext) Step 2
            temp1 = Hex2Dez(Mid$(Ciphertext, i, 2))
            temp2 = Hex2Dez(Mid$(Key, i, 2))
            plaintext = plaintext + Chr(temp1 Xor temp2)
        Next i


        Screensavepwd = plaintext
    Else
        Screensavepwd = " No Password"
    End If


End Function


Function Hex2Dez&(H$)


    If Left$(H$, 2) <> "&H" Then
        H$ = "&H" + H$
    End If


    Hex2Dez& = Val(H$)
End Function



Private Sub cpy1_Click()
Clipboard.Clear
Clipboard.SetText Text1.Text
End Sub

Private Sub Form_Load()
SetLoaded
Text1 = "This Program Has Been Loaded " & GetLoaded & " Times."
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Form1.PopupMenu Form1.opt1, 1
    End If
End Sub

Private Sub sv1_Click()
On Error Resume Next
cd1.Filter = "Screen Saver pwd(*.txt)|*.txt"
cd1.ShowOpen
If cd1.filename = "" Then Exit Sub
Open cd1.filename For Output As #1
Print #1, Text1.Text
Close #1
Form1.Show
End Sub
