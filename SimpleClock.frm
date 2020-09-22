VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Simple Clock"
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "SimpleClock.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "SimpleClock.frx":08CA
   ScaleHeight     =   855
   ScaleWidth      =   3630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1680
      Top             =   0
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   3240
      Picture         =   "SimpleClock.frx":107C
      Top             =   45
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   2880
      Picture         =   "SimpleClock.frx":1386
      Top             =   45
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "DigitMed"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   120
      Picture         =   "SimpleClock.frx":1690
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Simple Desktop Clock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   0
      Picture         =   "SimpleClock.frx":1AD2
      Top             =   960
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   0
      Picture         =   "SimpleClock.frx":1ECC
      Top             =   840
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   0
      Picture         =   "SimpleClock.frx":22C6
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim user As Integer
'Variables
'Switch to turn drah on and off.
Dim MoveScreen As Boolean
'Vars to get the mouse position on the form.
'   you are draging.
Dim MousA As Integer
Dim MousB As Integer
'Vars for moving the form.
Dim CurrX As Integer
Dim CurrY As Integer

Private Sub Image4_Click()
Form2.PopupMenu Form2.file, 1

End Sub
Sub pause(time)
 Start = time
 Do: DoEvents: Loop While Timer - Start <= time
End Sub

Private Sub Image5_Click()
Form1.WindowState = vbMinimized

End Sub

Private Sub Image6_Click()
Unload Me
End

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = Image2.Picture
MoveScreen = True ' make form movable while the mouse is down.
    'Get the initial coordinates of the mouse on the form.
    MousA = X
    MousB = Y

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If the mouse is down on the form then.
    If MoveScreen Then
        'Calculate the new x,y position for the form.
        '   NB. This is dependant on the X and Y vars on the Form_MouseMove,
        '   you can use objects MouseMove function. i.e. a Label or Textbox.
        CurrX = Form1.Left - MousA + X
        CurrY = Form1.Top - MousB + Y

        'Move the form to the new X,Y.
        Form1.Move CurrX, CurrY     ' move form.
        End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = Image3.Picture
MoveScreen = False

End Sub

Private Sub Timer1_Timer()
Label2.Caption = time & " ; " & Month(Now) & "/" & Day(Now) & "/" & Year(Now)
'You could also us: Label2.caption = Hour(now) &":" & minute(now) & ":" & second(now)
'But the time code is easier and more efficient!
If Form2.Text1.Text = Label2.Caption Then
Beep
pause 1
Beep
pause 1
Beep
pause 1
Beep
pause 1
Beep
pause 1
Beep
pause 1
Beep
pause 1
Beep
MsgBox "DING DING DING! ALARM!", vbInformation
End If

End Sub
