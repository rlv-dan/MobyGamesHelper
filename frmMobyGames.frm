VERSION 5.00
Begin VB.Form frmMobyGames 
   Caption         =   "MobyGames Credits Helper"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   Icon            =   "frmMobyGames.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkOnTop 
      Caption         =   "Always On Top"
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5280
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.VScrollBar vScrollBatchSend 
      Height          =   255
      LargeChange     =   2
      Left            =   5760
      Max             =   20
      Min             =   1
      SmallChange     =   2
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5160
      Value           =   1
      Width           =   255
   End
   Begin VB.TextBox txtBatchSend 
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   5160
      Width           =   495
   End
   Begin VB.CheckBox chkCapitalizeNames 
      Caption         =   "Capitalize names"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CheckBox chkCapitalizeTitles 
      Caption         =   "Capitalize titles"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4800
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load from Clipboard"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "p"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<<"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7680
      Top             =   0
   End
   Begin VB.ListBox lstCredits 
      Height          =   4155
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   7935
   End
   Begin VB.CheckBox chkSendKeys 
      Caption         =   "Send Ctrl+V and Tab after pressing Pause or F12"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4920
      Value           =   1  'Checked
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "times"
      Height          =   255
      Left            =   6120
      TabIndex        =   12
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Repeat"
      Height          =   255
      Left            =   4605
      TabIndex        =   11
      Top             =   5190
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmMobyGames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sCurrentLine As String
Dim iCurrentLine As Integer
Dim bProgramClick As Boolean
Dim iRepeat As Integer

Private Sub Form_Load()

    MakeTopMost (frmMobyGames.hWnd)

    vScrollBatchSend.Value = vScrollBatchSend.Max
    iRepeat = -1

    iCurrentLine = 0

    lstCredits.AddItem ("How to use:")
    lstCredits.AddItem ("   Prepare credits in your normal text editor (I recommend NotePad2) according to the format listed below.")
    lstCredits.AddItem ("   Load credits with the button below. Press the arrow button to minimize the window.")
    lstCredits.AddItem ("   The text will be automatically formatted, capitalized and cleaned up.")
    lstCredits.AddItem ("   Start the credits wizard and go to the groups page, or skip directly to add if there are no groups.")
    lstCredits.AddItem ("   Set the cursor in the input box and press PAUSE (alt: F12). Repeat until all groups have been added.")
    lstCredits.AddItem ("   Continue to the add page and do the same. All you have to do to add credits is to press pause!")
    lstCredits.AddItem ("   The program works by copying each line to the clpbioard and sending ctrl+v and tab to the browser.")
    lstCredits.AddItem ("   Please pause while names are being looked-up, so you don't overload the server with requests!")
    lstCredits.AddItem ("   Hotkeys: [SCRLCK] = back , [PAUSE] = forward , [F12] = forward, same as pause")
    lstCredits.AddItem ("   This program is not perfect. You must still be very observant to make sure that no mistakes occur!")
    lstCredits.AddItem ("   ")
    lstCredits.AddItem ("   Assumed Format:")
    lstCredits.AddItem ("")
    lstCredits.AddItem ("      [Company 1]")
    lstCredits.AddItem ("")
    lstCredits.AddItem ("      Role 1")
    lstCredits.AddItem ("      Name 1")
    lstCredits.AddItem ("      Name 2")
    lstCredits.AddItem ("      Name 3")
    lstCredits.AddItem ("      ...")
    lstCredits.AddItem ("")
    lstCredits.AddItem ("      Role 2")
    lstCredits.AddItem ("      Name 1 , Name 2 , Name 3")
    lstCredits.AddItem ("      Name 4")
    lstCredits.AddItem ("      FirstName 'Nickname' LastName 5")
    lstCredits.AddItem ("      ...")
    lstCredits.AddItem ("")
    lstCredits.AddItem ("      [Company 2]")
    lstCredits.AddItem ("")
    lstCredits.AddItem ("      Role 1")
    lstCredits.AddItem ("      ...")
    
End Sub

Private Sub chkOnTop_Click()

    If chkOnTop.Value = 1 Then MakeTopMost (Me.hWnd) Else MakeNormal (Me.hWnd)

End Sub

Private Sub cmdLoad_Click()

    lstCredits.Clear
    lstCredits.AddItem ("")
    Dim bPrevWasEmpty As Boolean
    txt = Split(Clipboard.GetText, vbNewLine)
    For n = LBound(txt) To UBound(txt)
        tmp = Split(txt(n), ",")
        txt(n) = ""
        For nn = LBound(tmp) To UBound(tmp)
            If bPrevWasEmpty = False Then
                tmp(nn) = FixAlias1(tmp(nn))
                tmp(nn) = FixAlias2(tmp(nn))
                tmp(nn) = FixAlias3(tmp(nn))
            End If
            txt(n) = txt(n) & ", " & Trim(tmp(nn))
        Next
        lstCredits.AddItem Mid(txt(n), 2)
        If Trim(txt(n)) = "" Then bPrevWasEmpty = True Else bPrevWasEmpty = False
    Next
    lstCredits.AddItem ("")
    
    iCurrentLine = 0

    Call format_credits

End Sub

Private Sub cmdNext_Click()
    
nxt:
    iCurrentLine = iCurrentLine + 1
    If iCurrentLine > lstCredits.ListCount - 1 Then iCurrentLine = lstCredits.ListCount - 1
    
    'skip empty lines
    tmp = lstCredits.List(iCurrentLine)
    If Trim(tmp) = "" And iCurrentLine < lstCredits.ListCount - 1 Then
        GoTo nxt
    End If

    GetLine
End Sub

Private Sub cmdPrev_Click()
    iCurrentLine = iCurrentLine - 1
    If iCurrentLine < 0 Then iCurrentLine = 0
    GetLine
End Sub

Private Sub Command3_Click()

    If Command3.Caption = "q" Then
        Command3.Caption = "p"
        Me.Height = 5955
    Else
        Command3.Caption = "q"
        Me.Height = 1000
    End If

End Sub


Private Sub Form_Resize()

    On Error Resume Next

    If Command3.Caption = "q" Then
    Else
        chkSendKeys.Top = frmMobyGames.ScaleHeight - 1020 + 405
        chkCapitalizeTitles.Top = frmMobyGames.ScaleHeight - 1140 + 405
        chkCapitalizeNames.Top = frmMobyGames.ScaleHeight - 1140 + 240 + 405
        chkOnTop.Top = frmMobyGames.ScaleHeight - 1140 + 240 + 240 + 405
        txtBatchSend.Top = frmMobyGames.ScaleHeight - 780 + 405
        vScrollBatchSend.Top = frmMobyGames.ScaleHeight - 780 + 405
        Label2.Top = frmMobyGames.ScaleHeight - 740 + 405
        Label3.Top = frmMobyGames.ScaleHeight - 740 + 405
        cmdLoad.Top = frmMobyGames.ScaleHeight - 1000 + 405
        cmdFormat.Top = frmMobyGames.ScaleHeight - 1000 + 405
        lstCredits.Height = frmMobyGames.ScaleHeight - 1700 + 405
        lstCredits.Width = frmMobyGames.ScaleWidth - 350 + 120
    
    End If

End Sub

Private Sub lstCredits_Click()

    If bProgramClick = True Then Exit Sub

    iCurrentLine = lstCredits.ListIndex
    GetLine

End Sub

Private Sub Timer1_Timer()

    'keycodes: http://msdn.microsoft.com/en-us/library/ms645540%28VS.85%29.aspx
    
    ret = GetAsyncKeyState(145)    'scrolllock
    If ret <> 0 Then
        Call cmdPrev_Click
    End If
    
    ret1 = GetAsyncKeyState(19)  'pause
    ret2 = GetAsyncKeyState(123)  'F12
    If ret1 <> 0 Or ret2 <> 0 Or iRepeat >= 0 Then
    
        If iRepeat = -1 Then
            iRepeat = Val(txtBatchSend.Text) - 1
        End If
        
        If iRepeat >= 0 Then iRepeat = iRepeat - 1
        
        Call cmdNext_Click
        If chkSendKeys.Value = 1 And (Left(sCurrentLine, 3) <> ">>>") Then
            Sleep (250)
            Call SendKeys("^v", True)
            Sleep (250)
            Call SendKeys("{TAB}", True)
            Sleep (250)
        Else
            iRepeat = -1
            Beep
        End If
    End If


End Sub

Private Sub GetLine()
    
    bProgramClick = True
    
    lstCredits.ListIndex = iCurrentLine
    sCurrentLine = lstCredits.List(iCurrentLine)
    Clipboard.Clear
    Label1.Caption = sCurrentLine
    If (Left(sCurrentLine, 3) = ">>>") Then
        Label1.BackColor = &HC0C0FF
    Else
        Label1.BackColor = &H80000018
    End If
    Label1 = Replace(Label1, "&", "&&")
    Clipboard.SetText sCurrentLine
    
    If iCurrentLine = lstCredits.ListCount - 1 Then 'EOF
        Label1.BackColor = &HC0FFC0
        Beep
    End If
    
    bProgramClick = False
    
End Sub

Private Sub format_credits()

    bProgramClick = True

    Dim bPreviousWasEmpty As Boolean
    bPreviousWasEmpty = False
    
    Dim sNames As String
    sNames = ""
    
    Dim tmp As String
    
    iFirstLine = 0

    Dim sCompanies(64) As String
    nCompanies = 0

    lastRole = ""

    For n = 0 To lstCredits.ListCount - 1
        
        tmp = lstCredits.List(n)
        If Trim(tmp) = "" Then
            bPreviousWasEmpty = True
            If sNames <> "" Then
                sNames = Replace(sNames, " + ", " , ")
                If Len(sNames) > 1000 Then
                    'MsgBox "Warning: line too long! Please limit each role to max 1000 chars. Split into several sections with the same role name. (Do not continue until this message goes away!)", vbCritical
                    lstCredits.Clear
                    lstCredits.AddItem ("Warning: too much name data in this chunk:")
                    lstCredits.AddItem ("")
                    lstCredits.AddItem (lastRole)
                    lstCredits.AddItem (sNames)
                    lstCredits.AddItem ("")
                    lstCredits.AddItem ("---------------------------------------------------------------------------------------------------------------------------------------------------------------")
                    lstCredits.AddItem ("")
                    lstCredits.AddItem ("Please limit each chunk of names to max 1000 chars.")
                    lstCredits.AddItem ("You can split a chunk into several parts with the same role name, like this:")
                    lstCredits.AddItem ("")
                    lstCredits.AddItem ("Role XYZ")
                    lstCredits.AddItem ("Name 1, Name 2, ... lot of names ..., Name 999      <-- more than 1000 characters!")
                    lstCredits.AddItem ("")
                    lstCredits.AddItem ("--> split -->")
                    lstCredits.AddItem ("")
                    lstCredits.AddItem ("Role XYZ")
                    lstCredits.AddItem ("Name 1, Name 2, Name 3")
                    lstCredits.AddItem ("")
                    lstCredits.AddItem ("Role XYZ")
                    lstCredits.AddItem ("Name 4, Name 5, Name 6")
                    lstCredits.AddItem ("...")
                    lstCredits.AddItem ("Role XYZ")
                    lstCredits.AddItem ("Name 998, Name 999")

                    Exit Sub
                End If
                lstCredits.List(iFirstLine) = sNames
            Else
                If (Trim(lstCredits.List(n - 1)) <> "") Then
                    sCompanies(nCompanies) = lstCredits.List(n - 1)
                    nCompanies = nCompanies + 1
                    
                    lstCredits.List(n - 1) = ">>> " & lstCredits.List(n - 1)
                End If
            End If
            sNames = ""
        Else
            If bPreviousWasEmpty = True Then    'roller
                bPreviousWasEmpty = False
                
                tmp = Replace(tmp, "“", Chr(34))
                tmp = Replace(tmp, "”", Chr(34))
                tmp = Replace(tmp, "`", "'")
                tmp = Replace(tmp, "´", "'")
                tmp = Replace(tmp, "’", "'")
                tmp = Replace(tmp, "‘", "'")
                tmp = Replace(tmp, """", "'")
                
                tmp = Replace(tmp, "/", " / ")
                tmp = Replace(tmp, "  ", " ")
                tmp = Replace(tmp, "F / X", "F/X")
                
                tmp = Trim(tmp)
                If chkCapitalizeTitles.Value = 1 Then tmp = CapitalizeAll(tmp)
                lstCredits.List(n) = tmp
                lastRole = tmp
            Else
                If sNames <> "" Then    'resterande namn
                    sNames = Replace(sNames, ", Inc", " Inc")
                    sNames = Replace(sNames, ", LLC", " LLC")
                    sNames = Replace(sNames, ", PLC", " PLC")
                    sNames = sNames & " + "
                Else    'första namnet
                    iFirstLine = n
                End If
                
                tmp = Replace(tmp, "“", Chr(34))
                tmp = Replace(tmp, "”", Chr(34))
                tmp = Replace(tmp, "`", "'")
                tmp = Replace(tmp, "´", "'")
                tmp = Replace(tmp, "’", "'")
                tmp = Replace(tmp, "‘", "'")
                tmp = Replace(tmp, """", "'")
                
                tmp = Trim(tmp)
                If chkCapitalizeNames.Value = 1 Then tmp = CapitalizeAll(tmp)
                sNames = sNames & tmp
                lstCredits.List(n) = ""
            End If
            
        End If
        
    Next
    
    lstCredits.AddItem (">>> EOF")

    'lägg till företagen längst upp
    For n = 0 To nCompanies - 1
        Call lstCredits.AddItem(sCompanies(n), n)
    Next
    Call lstCredits.AddItem("", 0)
    lstCredits.ListIndex = 0

restart:
    For n = 0 To lstCredits.ListCount - 2
        If Trim(lstCredits.List(n)) = "" And Trim(lstCredits.List(n + 1)) = "" Then
            lstCredits.RemoveItem (n)
            GoTo restart
        End If
    Next

    bProgramClick = False

End Sub

Function FixAlias1(ByVal sName As String) As String

    FixAlias1 = sName

    pos1 = InStr(sName, "'")
    If pos1 = 0 Then Exit Function
    pos2 = InStr(pos1 + 1, sName, "'")
    If pos2 = 0 Then Exit Function

    sName = Mid(sName, 1, pos1 - 1) & Mid(sName, pos2 + 1) & " ('" & Mid(sName, pos1 + 1, pos2 - pos1 - 1) & "')"
    FixAlias1 = Replace(sName, "  ", " ")
    FixAlias1 = Replace(sName, "()", "")

End Function

Function FixAlias2(ByVal sName As String) As String

    FixAlias2 = sName

    pos1 = InStr(sName, Chr(34))
    If pos1 = 0 Then Exit Function
    pos2 = InStr(pos1 + 1, sName, Chr(34))
    If pos2 = 0 Then Exit Function

    sName = Mid(sName, 1, pos1 - 1) & Mid(sName, pos2 + 1) & " ('" & Mid(sName, pos1 + 1, pos2 - pos1 - 1) & "')"
    FixAlias2 = Replace(sName, "  ", " ")
    FixAlias2 = Replace(sName, "()", "")

End Function

Function FixAlias3(ByVal sName As String) As String

    FixAlias3 = sName

    pos1 = InStr(sName, "(")
    If pos1 = 0 Then Exit Function
    pos2 = InStr(pos1 + 1, sName, ")")
    If pos2 = 0 Then Exit Function

    sName = Mid(sName, 1, pos1 - 1) & Mid(sName, pos2 + 1) & " (" & Mid(sName, pos1 + 1, pos2 - pos1 - 1) & ")"
    FixAlias3 = Replace(sName, "  ", " ")

End Function

Function UppercaseWord(sText As String, sWord As String) As String

    UppercaseWord = sText
    
    If UCase(sText) = UCase(sWord) Then
        UppercaseWord = UCase(sWord)
        Exit Function
    End If

    pos1 = InStr(UCase(sText), " " & UCase(sWord) & " ")
    If pos1 > 0 Then
        UppercaseWord = Mid(sText, 1, pos1 - 1) & " " & UCase(sWord) & " " & Mid(sText, pos1 + Len(sWord) + 2)
        Exit Function
    End If

    If UCase(Left(sText, Len(sWord) + 1)) = UCase(sWord & " ") Then
        UppercaseWord = UCase(sWord) & Mid(sText, Len(sWord) + 1)
        Exit Function
    End If

    If UCase(Left(sText, Len(sWord) + 1)) = UCase(sWord & "-") Then
        UppercaseWord = UCase(sWord) & "-" & Mid(sText, Len(sWord) + 2)
        Exit Function
    End If
    
    If UCase(Right(sText, Len(sWord) + 1)) = UCase(" " & sWord) Then
        UppercaseWord = Left(sText, Len(sText) - Len(sWord) - 1) & " " & UCase(sWord)
        Exit Function
    End If


    pos1 = InStr(1, sText, "-" & sWord & "-", vbTextCompare)
    If pos1 > 0 Then
        UppercaseWord = Mid(sText, 1, pos1 - 1) & "-" & UCase(sWord) & "-" & Mid(sText, pos1 + Len(sWord) + 2)
        Exit Function
    End If
'QA Technical Requirements Group (Trg)
    pos1 = InStr(1, sText, "-" & sWord & " ", vbTextCompare)
    If pos1 > 0 Then
        UppercaseWord = Mid(sText, 1, pos1 - 1) & "-" & UCase(sWord) & " " & Mid(sText, pos1 + Len(sWord) + 2)
        Exit Function
    End If

    pos1 = InStr(1, sText, " " & sWord & "-", vbTextCompare)
    If pos1 > 0 Then
        UppercaseWord = Mid(sText, 1, pos1 - 1) & " " & UCase(sWord) & "-" & Mid(sText, pos1 + Len(sWord) + 2)
        Exit Function
    End If


    pos1 = InStr(1, sText, "(" & sWord & ")", vbTextCompare)
    If pos1 > 0 Then
        UppercaseWord = Mid(sText, 1, pos1 - 1) & "(" & UCase(sWord) & ")" & Mid(sText, pos1 + Len(sWord) + 2)
        Exit Function
    End If

    pos1 = InStr(1, sText, "(" & sWord & " ", vbTextCompare)
    If pos1 > 0 Then
        UppercaseWord = Mid(sText, 1, pos1 - 1) & "(" & UCase(sWord) & " " & Mid(sText, pos1 + Len(sWord) + 2)
        Exit Function
    End If

    pos1 = InStr(1, sText, " " & sWord & ")", vbTextCompare)
    If pos1 > 0 Then
        UppercaseWord = Mid(sText, 1, pos1 - 1) & " " & UCase(sWord) & ")" & Mid(sText, pos1 + Len(sWord) + 2)
        Exit Function
    End If

    If UCase(Left(sText, Len(sWord) + 1)) = UCase(sWord & "-") Then
        UppercaseWord = UCase(sWord) & "-" & Mid(sText, Len(sWord) + 2)
        Exit Function
    End If

    If UCase(Right(sText, Len(sWord) + 1)) = UCase("-" & sWord) Then
        UppercaseWord = Left(sText, Len(sText) - Len(sWord) - 1) & "-" & UCase(sWord)
        Exit Function
    End If

End Function

Function LowercaseWord(sText As String, sWord As String) As String

    LowercaseWord = sText
    
    If LCase(sText) = LCase(sWord) Then
        LowercaseWord = LCase(sWord)
        Exit Function
    End If

    pos1 = InStr(LCase(sText), " " & LCase(sWord) & " ")
    If pos1 > 1 Then
        LowercaseWord = Mid(sText, 1, pos1 - 1) & " " & LCase(sWord) & " " & Mid(sText, pos1 + Len(sWord) + 2)
        Exit Function
    End If

    If LCase(Right(sText, Len(sWord) + 1)) = LCase(" " & sWord) Then
        LowercaseWord = Left(sText, Len(sText) - Len(sWord) - 1) & " " & LCase(sWord)
        Exit Function
    End If

End Function

Public Function CapitalizeAll(tmp As String) As String
    
    tmp = Capitalize(tmp)
    
    tmp = UppercaseWord(tmp, "QA")
    tmp = UppercaseWord(tmp, "CS")
    tmp = UppercaseWord(tmp, "IT")
    tmp = UppercaseWord(tmp, "VP")
    tmp = UppercaseWord(tmp, "EVP")
    tmp = UppercaseWord(tmp, "EVP,")
    tmp = UppercaseWord(tmp, "VP,")
    tmp = UppercaseWord(tmp, "GM")
    tmp = UppercaseWord(tmp, "HR")
    tmp = UppercaseWord(tmp, "CEO")
    tmp = UppercaseWord(tmp, "CEO,")
    tmp = UppercaseWord(tmp, "COO")
    tmp = UppercaseWord(tmp, "COO,")
    tmp = UppercaseWord(tmp, "CCO")
    tmp = UppercaseWord(tmp, "CCO,")
    tmp = UppercaseWord(tmp, "CTO")
    tmp = UppercaseWord(tmp, "GUI")
    tmp = UppercaseWord(tmp, "LLP")
    tmp = UppercaseWord(tmp, "TRG")
    tmp = UppercaseWord(tmp, "MIS")
    tmp = UppercaseWord(tmp, "DBA")
    tmp = UppercaseWord(tmp, "FX")
    tmp = UppercaseWord(tmp, "FMV")
    tmp = UppercaseWord(tmp, "2D")
    tmp = UppercaseWord(tmp, "3D")
    tmp = UppercaseWord(tmp, "CG")
    tmp = UppercaseWord(tmp, "PC")
    tmp = UppercaseWord(tmp, "PR")
    tmp = UppercaseWord(tmp, "SVP")
    tmp = UppercaseWord(tmp, "SVP,")
    tmp = UppercaseWord(tmp, "AI")
    tmp = UppercaseWord(tmp, "VO")
    tmp = UppercaseWord(tmp, "TV")
    tmp = UppercaseWord(tmp, "PSX")
    tmp = UppercaseWord(tmp, "SFX")
    tmp = UppercaseWord(tmp, "UK")
    tmp = UppercaseWord(tmp, "US")
    tmp = UppercaseWord(tmp, "CGI")
    tmp = UppercaseWord(tmp, "UI")
    tmp = UppercaseWord(tmp, "CFO")
    tmp = UppercaseWord(tmp, "EMEA")
    tmp = UppercaseWord(tmp, "AV")
    tmp = UppercaseWord(tmp, "MPL")
    tmp = UppercaseWord(tmp, "QA-MPL")
    tmp = UppercaseWord(tmp, "QA-CL")
    tmp = UppercaseWord(tmp, "CRG")
    tmp = UppercaseWord(tmp, "QA-CRG")
    tmp = UppercaseWord(tmp, "MIS")
    tmp = UppercaseWord(tmp, "QA-MIS")
    tmp = UppercaseWord(tmp, "QA-AVL")
    tmp = UppercaseWord(tmp, "DBA")
    tmp = UppercaseWord(tmp, "DBS")
    tmp = UppercaseWord(tmp, "QA-DBA")
    tmp = UppercaseWord(tmp, "QC")
    tmp = UppercaseWord(tmp, "BGM")
    tmp = UppercaseWord(tmp, "CSQA")
    
    tmp = LowercaseWord(tmp, "of")
    tmp = LowercaseWord(tmp, "the")
    tmp = LowercaseWord(tmp, "to")
    tmp = LowercaseWord(tmp, "by")
    tmp = LowercaseWord(tmp, "and")
    tmp = LowercaseWord(tmp, "from")
    tmp = LowercaseWord(tmp, "at")
    tmp = LowercaseWord(tmp, "an")

    CapitalizeAll = tmp

End Function

Private Sub vScrollBatchSend_Change()
    If vScrollBatchSend = vScrollBatchSend.Max Then
        vScrollBatchSend.SmallChange = 1
        vScrollBatchSend.LargeChange = 1
    Else
        vScrollBatchSend.SmallChange = 2
        vScrollBatchSend.LargeChange = 2
    End If
    txtBatchSend = vScrollBatchSend.Max + 1 - vScrollBatchSend.Value
End Sub

