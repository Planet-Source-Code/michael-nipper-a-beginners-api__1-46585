Attribute VB_Name = "Module1"
'Ah API!  I know it looks difficult, but no worries.  This is not conceptual at all
'and once you memorize these 4 lines you'll be well on your way to bossing other programs
'around.  Well, the truth is, there really is no need to memorize them at all.  You can
'refer to the API Viewer that came with your Visual Basic program.  Once you open it
'goto file then open text file then Win32API.  Then, as an example, you can type in
'FindWindow and will see exactly what is written below.

'Or you could just load this Module or any other similar module(hopefully with more
'in it) into your project when you want to make your own program.
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpwindowname As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

'Constants that make it easier to program once you get going.  For instance it is easier
'to have to look up that that code to set text is &HC once and then make it a constant
'than to have to continuously check back in your code or look it up again.  So now,
'WM_SETTEXT and &HC and equal so instead of having to type out &HC you can type out
'WM_SETTEXT anywhere throughout your project.  If you wanted WM_SETTEXT to equal &HC
'only in this module instead of throughout the entire project, you could type out
'Private Const WM_SETTEXT = &HC instead of Public Const WM_SETTEXT = &HC which makes
'them equal throughout the entire project.
Public Const WM_SETTEXT = &HC
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101

'The Constant for he code for the Space bar.  Used the press buttons.
Public Const VK_SPACE = &H20












'**************************************************************************************
'Answers for Form2:
'**************************************************************************************

'Private Sub Command13_Click()
'Dim Calc As Long
'Dim CalcBtn As Long
'Calc = FindWindow("scicalc", vbNullString)
'CalcBtn = FindWindowEx(Calc, 0&, "button", "*")
'Call SendMessageLong(CalcBtn, WM_KEYDOWN, VK_SPACE, 0&)
'Call SendMessageLong(CalcBtn, WM_KEYUP, VK_SPACE, 0&)
'End Sub

'Private Sub Command14_Click()
'Dim Calc As Long
'Dim CalcBtn As Long
'Calc = FindWindow("scicalc", vbNullString)
'CalcBtn = FindWindowEx(Calc, 0&, "button", "/")
'Call SendMessageLong(CalcBtn, WM_KEYDOWN, VK_SPACE, 0&)
'Call SendMessageLong(CalcBtn, WM_KEYUP, VK_SPACE, 0&)
'End Sub

'Private Sub Command15_Click()
'Dim Calc As Long
'Dim CalcBtn As Long
'Calc = FindWindow("scicalc", vbNullString)
'CalcBtn = FindWindowEx(Calc, 0&, "button", "-")
'Call SendMessageLong(CalcBtn, WM_KEYDOWN, VK_SPACE, 0&)
'Call SendMessageLong(CalcBtn, WM_KEYUP, VK_SPACE, 0&)
'End Sub

'Private Sub Command16_Click()
'Dim Calc As Long
'Dim CalcBtn As Long
'Calc = FindWindow("scicalc", vbNullString)
'CalcBtn = FindWindowEx(Calc, 0&, "button", "C")
'Call SendMessageLong(CalcBtn, WM_KEYDOWN, VK_SPACE, 0&)
'Call SendMessageLong(CalcBtn, WM_KEYUP, VK_SPACE, 0&)
'End Sub
