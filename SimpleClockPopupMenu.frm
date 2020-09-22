VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu SetAlarm 
         Caption         =   "Set Alarm"
      End
      Begin VB.Menu stoparlam 
         Caption         =   "Stop Alarm"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SetAlarm_Click()
AskTime = InputBox("Enter Time, With format: 12:00:10 AM ; 12/9/2000")
Text1.Text = AskTime

End Sub

Private Sub stoparlam_Click()
Text1.Text = ""

End Sub
