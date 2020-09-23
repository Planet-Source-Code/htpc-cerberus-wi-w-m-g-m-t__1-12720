VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WiÑÐØw§ MÅÑÅgèmèÑt"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2760
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   2760
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   4920
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4920
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Left            =   4920
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Minimize"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label12 
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   2280
      Width           =   7815
   End
   Begin VB.Label Label11 
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "®è§tá®t"
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "§hut ÐØwÑ"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "M§ PâìñT"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "ÇãLÇulàtØ®"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "ñötèPâÐ"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Çlö§è ÇÐ Plàyè®"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Øpèñ ÇÐ Plàyè®"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "IñTe®Ñèt èxpLØ®è®"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "       ßà§ìÇ ØptiØÑ§"
      Height          =   2055
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "This Was Made By [H]tpc [C]erberus", 0, "(c) 2000-2001"
Unload Me
End

End Sub

Private Sub Command2_Click()
Me.WindowState = 1

End Sub


Private Sub Form_Load()
Timer1.Interval = 1000
StayOnTop Form1, True
Timer2.Enabled = True
Let Timer2.Enabled = True
    Let Timer2.Interval = "20"
    Let Label12.Height = "255"
    Let Label12.Width = "7700"
    Let Label12.Caption = "Full Version With New Options && Free Updates Email Me @ Nienleiterhosen@aol.com Registration Is $ 5"
    Let Label12.Font = "lcdD"
    Let Label12.ForeColor = vbRed
    Let Label12.BackColor = vbBlack
    Let Label12.FontSize = "10"
    Let Label12.FontBold = False
    



    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbBlack
    Label1.FontItalic = False
Label2.ForeColor = vbBlack
    Label2.FontItalic = False
    Label3.ForeColor = vbBlack
    Label3.FontItalic = False
    Label4.ForeColor = vbBlack
    Label4.FontItalic = False
    Label5.ForeColor = vbBlack
    Label5.FontItalic = False
    Label6.ForeColor = vbBlack
    Label6.FontItalic = False
    Label7.ForeColor = vbBlack
    Label7.FontItalic = False
    Label8.ForeColor = vbBlack
    Label8.FontItalic = False
    Label9.ForeColor = vbBlack
    Label9.FontItalic = False
    
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False


End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbRed
    Label1.FontItalic = True

    
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True



End Sub

Private Sub Label10_Click()
Label10.Caption = Time
End Sub

Private Sub Label12_Click()
MsgBox "The Full Version Will Contain Over 20 Options With Free Updates!", 0, "Enjoy This Demo"
End Sub

Private Sub Label2_Click()
Dim X As Variant
X = Shell("C:\Program Files\Internet Explorer\Iexplore.exe")

End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.FontBold = True
Label2.ForeColor = vbGreen

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlue
    Label2.FontItalic = True
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.FontBold = False
Label2.ForeColor = vbBlue

End Sub

Private Sub Label3_Click()
On Error Resume Next
    retvalue = mciSendString("set CDAudio door open", returnstring, 127, 0)

End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontBold = True
Label3.ForeColor = vbGreen

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbBlue
    Label3.FontItalic = True
    Label2.ForeColor = vbBlack
    Label2.FontItalic = False
    
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontBold = True
Label3.ForeColor = vbBlue

End Sub

Private Sub Label4_Click()
On Error Resume Next
    retvalue = mciSendString("set CDAudio door closed", returnstring, 127, 0)

End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.FontBold = False
Label4.ForeColor = vbGreen

End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbBlue
    Label4.FontItalic = True
    Label3.ForeColor = vbBlack
    Label3.FontItalic = False
    
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbBlue
Label4.FontItalic = False
End Sub

Private Sub Label5_Click()
Dim LF As Variant
    LF = Shell(" notepad ", 3)

End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.FontBold = True
Label5.ForeColor = vbGreen

End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbBlue
    Label5.FontItalic = True
    Label4.ForeColor = vbBlack
    Label4.FontItalic = False
    
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.FontBold = False
Label5.ForeColor = vbBlue

End Sub

Private Sub Label6_Click()
Dim MyDick As Variant
MyDick = Shell("C:\windows\Calc.exe")

End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontBold = True
Label6.ForeColor = vbGreen

End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = vbBlue
    Label6.FontItalic = True
    Label5.ForeColor = vbBlack
    Label5.FontItalic = False
    
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontBold = False
Label6.ForeColor = vbBlue

End Sub

Private Sub Label7_Click()
Dim Y As Variant
Y = Shell("C:\Program Files\Accessories\Mspaint.exe")

End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.FontBold = True
Label7.ForeColor = vbGreen

End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = vbBlue
    Label7.FontItalic = True
    Label6.ForeColor = vbBlack
    Label6.FontItalic = False
    
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.FontBold = False
Label7.ForeColor = vbBlue

End Sub

Private Sub Label8_Click()
ExitWindows EWX_LogOff, &HFFFFFFFF
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.FontBold = True
Label8.ForeColor = vbGreen

End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = vbBlue
    Label8.FontItalic = True
    Label7.ForeColor = vbBlack
    Label7.FontItalic = False
    
End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.FontBold = False
Label8.ForeColor = vbBlue

End Sub

Private Sub Label9_Click()
Dim intRetVal As Integer
    intRetVal = ExitWindows(EW_RESTARTWINDOWS, 0)
    ExitWinRestart = intRetVal
End Sub


Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.FontBold = True
Label9.ForeColor = vbGreen

End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = vbBlue
    Label9.FontItalic = True
    Label8.ForeColor = vbBlack
    Label8.FontItalic = False
    
End Sub

Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.FontBold = False
Label9.ForeColor = vbBlue

End Sub

Private Sub Timer1_Timer()
Label10.Caption = Time
    Label11.Caption = Date
End Sub

Private Sub Timer2_Timer()
Label12.Move Label12.Left - 25


    If Label12.Left < -Label12.Width Then
        Label12.Left = Form1.Width
    End If

End Sub
