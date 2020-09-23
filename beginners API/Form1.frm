VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Notepad Form"
   ClientHeight    =   1965
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Change"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "What to change title to"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text to send"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Menu gethwnd 
      Caption         =   "Get hWnd..."
      Begin VB.Menu notepad 
         Caption         =   "Notepad"
      End
      Begin VB.Menu textbox 
         Caption         =   "Textbox"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim is what you use to declare variables.  We declare two variables, one for the main
'notepad form and one for the textbox that is on the form.  In order for us to find the
'textbox we must first find the main window it is on.
Dim notepad As Long
Dim NotePadTxt As Long

'This returns a value using the function FindWindow that we declared in Module1.bas
'What it returns is the handle of the window so that you can manipulate it.  Handles
'always change so it is not possible to just say find this hWnd(handle).  However,
'every window has a classname so that you can find the hWnd from that.  FindWindow was
'declared as a FUNCTION and therefore must return a value.  In this case the value is
' the handle of the main notepad window.  Functions always return values, while subs
'do not.
'variable = FindWindow(Class Name, Windowtext)
'In this case we do not want to use the windowtext to get the hWnd because it is possible
'to change the window text with this program and because notepad's caption changes
'anyway.  So therefore to tell visual basic that we do not want to use this declaration
'as an option, we use vbNullString.  We use vbNullString because that declaration was
'originally declared a string, if it were declare a long then we would want to set it to
'0& instead.
notepad = FindWindow("Notepad", vbNullString)
'If notepad = 0 then notepad is not loaded.  If it were loaded it would return some
'numerical value which would be its hWnd.
If notepad = 0 Then MsgBox "Please open notepad.", vbExclamation, "Error"
'This now finds the textbox on notepad.  It does this by using the main window's handle
'which we extracted and set equal to notepad a couple lines up.
NotePadTxt = FindWindowEx(notepad, 0, "edit", vbNullString)
'This calling the function we declared in Module1.bas which was called SendMessageByString
'and was given the declarations hwnd As Long, wMsg As Long, wParam As Long, lParam As
'String.  hWnd, the first declaration, is where you put the hWnd of the object in which
'you want to send text to.  In this case we have set the hWnd of the textbox on notepad
'to NotePadTxt by using the FindWindowEX function.  Also, we want it to send text to
'this textbox, so we set the next declaration to WM_SETTEXT which is a constant in
'Module1.bas.  The next declaration is not necessary in this operation so is set to 0
'since it is declared as a long in the module.  If it were declared as a string,
'vbNullString would substitute it.  The last declaration is where you put the text which
'you want to send to the object.  Since we want to send the text that is in text1 then
'we put Text1.text
Call SendMessageByString(NotePadTxt, WM_SETTEXT, 0, Text1.Text)
End Sub

Private Sub Command2_Click()
'Here we are again declaring variables like we did above.
Dim notepad As Long
Dim NotePadTxt As Long

'We set the variable notepad equal to the handle of the main notepad window according
'to the classname of the window.  To find the classname of a window, you can use the
'spy that came with your Visual Basic program or you can download one from the internet.
notepad = FindWindow("Notepad", vbNullString)
If notepad = 0 Then MsgBox "Please open notepad.", vbExclamation, "Error"
Call SendMessageByString(notepad, WM_SETTEXT, 0, Text2.Text)
End Sub

Private Sub Form_Load()
Form2.Show
End Sub

Private Sub notepad_Click()
'Declare variables
Dim notepad As Long
Dim NotePadTxt As Long

notepad = FindWindow("Notepad", vbNullString)
'If there is no handle for notepad then it is not open, in which case we open a msgbox
'informing the user to open notepad before proceeding.
If notepad = 0 Then
MsgBox "Please open notepad.", vbExclamation, "Error"
'After displaying the msgbox we exit sub to avoid it continuing and displaying a second
'message box displaying 0 which is what it returned for the handle of notepad since
'it was not open.
Exit Sub
End If

'Display the handle of notepad in a msgbox if it is open.
MsgBox notepad
End Sub

Private Sub textbox_Click()
Dim notepad As Long
Dim NotePadTxt As Long

notepad = FindWindow("Notepad", vbNullString)
If notepad = 0 Then
MsgBox "Please open notepad.", vbExclamation, "Error"
Exit Sub
End If
NotePadTxt = FindWindowEx(notepad, 0, "edit", vbNullString)
MsgBox NotePadTxt
End Sub
