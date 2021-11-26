Attribute VB_Name = "Module1"

Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
    
    Global Const SWP_NOMOVE = 2
    Global Const SWP_NOSIZE = 1
    Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Global Const HWND_TOPMOST = -1
    Global Const HWND_NOTOPMOST = -2
    
Declare Function SetWindowWord Lib "user32" (ByVal hwnd&, ByVal nIndex&, ByVal wNewWord&) As Long
Public lngOrigParenthWnd&


    

