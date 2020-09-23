VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Calculator Form"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2970
   LinkTopic       =   "Form2"
   ScaleHeight     =   4200
   ScaleWidth      =   2970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Now you try:"
      Height          =   1335
      Left            =   480
      TabIndex        =   12
      Top             =   2640
      Width           =   1935
      Begin VB.CommandButton Command16 
         Caption         =   "C"
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton Command15 
         Caption         =   "-"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton Command14 
         Caption         =   "/"
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command13 
         Caption         =   "*"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton Command12 
      Caption         =   "="
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "0"
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "+"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "3"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "2"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "6"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "5"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "1"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "9"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "8"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Declaring variables
Dim Calc As Long
Dim CalcBtn As Long

'Get the handle of the calculator using it's classname, scicalc.
Calc = FindWindow("scicalc", vbNullString)
'Now get the handle of the button on the calculator with the text "7" in it and the
'classname of button.
CalcBtn = FindWindowEx(Calc, 0&, "button", "7")
'SendMessageLong here is used to click the button.  First we use it to press the space
'bar down using the constants we declared in module1.bas, WM_KEYDOWN and VK_SPACE.
'WM_KEYDOWN is used  to press a button down and VK_SPACE is the code for the space
'bar.
Call SendMessageLong(CalcBtn, WM_KEYDOWN, VK_SPACE, 0&)
'After pressing the button down, you must then let it back up by using WM_KEYUP.
Call SendMessageLong(CalcBtn, WM_KEYUP, VK_SPACE, 0&)
End Sub

Private Sub Command10_Click()
Dim Calc As Long
Dim CalcBtn As Long

Calc = FindWindow("scicalc", vbNullString)
CalcBtn = FindWindowEx(Calc, 0&, "button", "+")
Call SendMessageLong(CalcBtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(CalcBtn, WM_KEYUP, VK_SPACE, 0&)
End Sub

Private Sub Command11_Click()
Dim Calc As Long
Dim CalcBtn As Long

Calc = FindWindow("scicalc", vbNullString)
CalcBtn = FindWindowEx(Calc, 0&, "button", "0")
Call SendMessageLong(CalcBtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(CalcBtn, WM_KEYUP, VK_SPACE, 0&)
End Sub

Private Sub Command12_Click()
Dim Calc As Long
Dim CalcBtn As Long

Calc = FindWindow("scicalc", vbNullString)
CalcBtn = FindWindowEx(Calc, 0&, "button", "=")
Call SendMessageLong(CalcBtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(CalcBtn, WM_KEYUP, VK_SPACE, 0&)
End Sub

Private Sub Command13_Click()
'This is where you try to go on your own and create your first program that can
'control other programs using API.  I have started you out...and if you get stuck
'and can't figure out what you did wrong, I posted the answers in module1.bas at the
'bottom!

Dim Calc As Long
Dim CalcBtn As Long

'Your Code Starts Here!  Delete the following line of code before proceeding!
MsgBox "You haven't edited this yet!", vbExclamation, "Sucker"
End Sub

Private Sub Command14_Click()
'This is where you try to go on your own and create your first program that can
'control other programs using API.  I have started you out...and if you get stuck
'and can't figure out what you did wrong, I posted the answers in module1.bas at the
'bottom!

Dim Calc As Long
Dim CalcBtn As Long

'Your Code Starts Here!  Delete the following line of code before proceeding!
MsgBox "You haven't edited this yet!", vbExclamation, "Sucker"
End Sub

Private Sub Command15_Click()
'This is where you try to go on your own and create your first program that can
'control other programs using API.  I have started you out...and if you get stuck
'and can't figure out what you did wrong, I posted the answers in module1.bas at the
'bottom!

Dim Calc As Long
Dim CalcBtn As Long

'Your Code Starts Here!  Delete the following line of code before proceeding!
MsgBox "You haven't edited this yet!", vbExclamation, "Sucker"
End Sub

Private Sub Command16_Click()
'This is where you try to go on your own and create your first program that can
'control other programs using API.  I have started you out...and if you get stuck
'and can't figure out what you did wrong, I posted the answers in module1.bas at the
'bottom!

Dim Calc As Long
Dim CalcBtn As Long

'Your Code Starts Here!  Delete the following line of code before proceeding!
MsgBox "You haven't edited this yet!", vbExclamation, "Sucker"
End Sub

Private Sub Command2_Click()
Dim Calc As Long
Dim CalcBtn As Long

Calc = FindWindow("scicalc", vbNullString)
CalcBtn = FindWindowEx(Calc, 0&, "button", "8")
Call SendMessageLong(CalcBtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(CalcBtn, WM_KEYUP, VK_SPACE, 0&)
End Sub

Private Sub Command3_Click()
Dim Calc As Long
Dim CalcBtn As Long

Calc = FindWindow("scicalc", vbNullString)
CalcBtn = FindWindowEx(Calc, 0&, "button", "9")
Call SendMessageLong(CalcBtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(CalcBtn, WM_KEYUP, VK_SPACE, 0&)
End Sub

Private Sub Command4_Click()
Dim Calc As Long
Dim CalcBtn As Long

Calc = FindWindow("scicalc", vbNullString)
CalcBtn = FindWindowEx(Calc, 0&, "button", "4")
Call SendMessageLong(CalcBtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(CalcBtn, WM_KEYUP, VK_SPACE, 0&)
End Sub

Private Sub Command5_Click()
Dim Calc As Long
Dim CalcBtn As Long

Calc = FindWindow("scicalc", vbNullString)
CalcBtn = FindWindowEx(Calc, 0&, "button", "1")
Call SendMessageLong(CalcBtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(CalcBtn, WM_KEYUP, VK_SPACE, 0&)
End Sub

Private Sub Command6_Click()
Dim Calc As Long
Dim CalcBtn As Long

Calc = FindWindow("scicalc", vbNullString)
CalcBtn = FindWindowEx(Calc, 0&, "button", "5")
Call SendMessageLong(CalcBtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(CalcBtn, WM_KEYUP, VK_SPACE, 0&)
End Sub

Private Sub Command7_Click()
Dim Calc As Long
Dim CalcBtn As Long

Calc = FindWindow("scicalc", vbNullString)
CalcBtn = FindWindowEx(Calc, 0&, "button", "6")
Call SendMessageLong(CalcBtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(CalcBtn, WM_KEYUP, VK_SPACE, 0&)
End Sub

Private Sub Command8_Click()
Dim Calc As Long
Dim CalcBtn As Long

Calc = FindWindow("scicalc", vbNullString)
CalcBtn = FindWindowEx(Calc, 0&, "button", "2")
Call SendMessageLong(CalcBtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(CalcBtn, WM_KEYUP, VK_SPACE, 0&)
End Sub

Private Sub Command9_Click()
Dim Calc As Long
Dim CalcBtn As Long

Calc = FindWindow("scicalc", vbNullString)
CalcBtn = FindWindowEx(Calc, 0&, "button", "3")
Call SendMessageLong(CalcBtn, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(CalcBtn, WM_KEYUP, VK_SPACE, 0&)
End Sub

