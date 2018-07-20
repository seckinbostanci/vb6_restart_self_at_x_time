VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zamanlý Kapanma"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTimeShower 
      Interval        =   100
      Left            =   2880
      Top             =   240
   End
   Begin VB.TextBox txtTimeView 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub Form_Load()

End Sub

Private Sub tmrTimeShower_Timer()
    
    Dim theHour As String
    Dim theMinute As String
    Dim theSecond As String
    Dim theMilisecond As String
    
    theHour = Hour(Now)
    theMinute = Minute(Now)
    theSecond = Second(Now)
    theMilisecond = Format$(Now(), "ms")

    txtTimeView.Text = theHour & ":" & theMinute & ":" & theSecond & ":" & theMilisecond

    If StrComp(theHour, "9", vbTextCompare) = 0 And _
       StrComp(theMinute, "0", vbTextCompare) = 0 And _
       StrComp(theSecond, "0", vbTextCompare) = 0 And _
       (StrComp(theMilisecond, "0", vbTextCompare) = 0 Or StrComp(theMilisecond, "0", vbTextCompare) < 101) Then
        Call ShellExecute(Me.hwnd, "Open", App.Path & "\" & App.EXEName & ".exe", "Restart", "", 1)
        Unload Form1
        Set Form1 = Nothing
    End If
    If StrComp(theHour, "18", vbTextCompare) = 0 And _
       StrComp(theMinute, "0", vbTextCompare) = 0 And _
       StrComp(theSecond, "0", vbTextCompare) = 0 Then
        Call ShellExecute(Me.hwnd, "Open", App.Path & "\" & App.EXEName & ".exe", "Restart", "", 1)
        Unload Form1
        Set Form1 = Nothing
    End If
End Sub
