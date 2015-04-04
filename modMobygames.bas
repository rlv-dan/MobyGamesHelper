Attribute VB_Name = "modMobygames"
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'API Sleep

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'api constants
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

'make a form normal
Public Sub MakeNormal(hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

'make a form topmost
Public Sub MakeTopMost(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
    
Public Sub TestPause()
    
    'codes: http://msdn.microsoft.com/en-us/library/ms645540(VS.85).aspx
    
    ret = GetAsyncKeyState(19)  'pause
    If ret <> 0 Then
        Debug.Print "HIT"
    End If

End Sub


Public Function Capitalize(FileName)

    New_Filename = ""
    Prev_Was_Space = True
    For num = 1 To Len(FileName)
    
        If Prev_Was_Space = True Then
            New_Filename = New_Filename & UCase$(Mid$(FileName, num, 1))
        Else
            New_Filename = New_Filename & LCase$(Mid$(FileName, num, 1))
        End If
        
        Prev_Was_Space = False

        Dim testChar As String
        
        If Mid$(FileName, num, 1) = " " Then Prev_Was_Space = True
        If Mid$(FileName, num, 1) = "(" Then Prev_Was_Space = True
        If Mid$(FileName, num, 1) = "[" Then Prev_Was_Space = True
        If Mid$(FileName, num, 1) = "-" Then Prev_Was_Space = True
        If Mid$(FileName, num, 1) = "." Then Prev_Was_Space = True
        If Mid$(FileName, num, 1) = "_" Then Prev_Was_Space = True
        If Mid$(FileName, num, 1) = Chr(34) Then Prev_Was_Space = True
        If num >= 2 Then
            If Mid$(FileName, num - 1, 2) = " '" Then Prev_Was_Space = True
            If Mid$(FileName, num - 1, 2) = " """ Then Prev_Was_Space = True
        End If

    Next
    
    Capitalize = New_Filename

End Function

