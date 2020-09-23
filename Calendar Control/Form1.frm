VERSION 5.00
Object = "*\A..\Calendar\Project1.vbp"
Begin VB.Form Form1 
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin Project1.RCalendar RCalendar1 
      Height          =   3735
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6588
      HeadingColor    =   -2147483630
      TitleColor      =   -2147483630
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MonthYearFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Calendar1_KeyPress(KeyAscii As Integer)
    MsgBox KeyAscii
End Sub

Private Sub UserControl11_KeyDown(KeyCode As Integer, Shift As Integer)
    MsgBox KeyCode
End Sub

Private Sub Command1_Click()
        RCalendar1.Value = Date
    End Sub

Private Sub Command2_Click()
UserControl11.daycolor = vbBlue
End Sub


Private Sub UserControl11_Click()

End Sub

