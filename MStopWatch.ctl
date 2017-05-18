VERSION 5.00
Begin VB.UserControl MStopWatch 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7710
   ScaleHeight     =   1080
   ScaleWidth      =   7710
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5760
      Top             =   120
   End
   Begin VB.Label lbl_Time 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1065
   End
End
Attribute VB_Name = "MStopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim seconds, minutes, hours As Long

Private Sub Timer1_Timer()
Clock_Ticker
End Sub

Private Sub UserControl_Initialize()
Reset
 UserControl.Height = lbl_Time.Height
 End Sub

Public Sub StartWatch()
 
Timer1.Enabled = True
End Sub

Public Sub StopWatch()
Timer1.Enabled = False
Reset
End Sub
Public Sub PAUSEWatch()
Timer1.Enabled = False
 
End Sub


Public Sub Clock_Ticker()
 seconds = seconds + 1
    If (seconds = 60) Then
        seconds = 0
        minutes = minutes + 1
        If (minutes = 60) Then
            minutes = 0
            hours = hours + 1
        End If
    End If
    UpdateLbl
End Sub

Public Sub UpdateLbl()
lbl_Time = hours & ":" & minutes & ":" & seconds
End Sub

Public Sub Reset()
minutes = 0
hours = 0
seconds = 0
lbl_Time = vbNullString
UpdateLbl
End Sub


Public Property Get Interval() As Variant
Interval = Timer1.Interval
End Property

Public Property Let Interval(ByVal vNewValue As Variant)
Timer1.Interval = vNewValue
PropertyChanged "Interval"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Timer1.Interval = PropBag.ReadProperty("Interval", Timer1.Interval)
  Set lbl_Time.Font = PropBag.ReadProperty("Font", lbl_Time.Font)
 

End Sub

Private Sub UserControl_Resize()
lbl_Time.Width = UserControl.Width
lbl_Time.Height = UserControl.Height

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Interval", UserControl.Timer1.Interval
 PropBag.WriteProperty "Font", UserControl.lbl_Time.Font, Ambient.Font
 



End Sub

 

Public Property Set Font(ByVal Nfont As StdFont)
 
  Set lbl_Time.Font = Nfont
    
Refresh
PropertyChanged "Font"

End Property

Public Property Get Font() As StdFont
Set Font = UserControl.lbl_Time.Font

End Property

 
 
