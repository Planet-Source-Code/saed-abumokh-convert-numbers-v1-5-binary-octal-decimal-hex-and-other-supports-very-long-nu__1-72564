Attribute VB_Name = "Other"
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub FlashWindow(Object As Object, FlashColor As Long, FlashTimes As Long, FlashTimeLimitMS As Long, NormalTimeLimitMS As Long, LastFlashTimeLimitMS As Long)
    Dim s1 As Long
    Dim i As Long
    Dim SaveBackColor As Long
    SaveBackColor = Object.BackColor
    
    For i = 1 To FlashTimes - 1
        s1 = GetTickCount
        Do
            DoEvents
            Object.BackColor = FlashColor
        Loop Until GetTickCount() - s1 > FlashTimeLimitMS
        
        s1 = GetTickCount
        Do
            DoEvents
            Object.BackColor = SaveBackColor
        Loop Until GetTickCount() - s1 > NormalTimeLimitMS
    Next
    
    s1 = GetTickCount
    Do
        DoEvents
        Object.BackColor = FlashColor
    Loop Until GetTickCount() - s1 > LastFlashTimeLimitMS
    
    Object.BackColor = SaveBackColor
End Sub

Public Sub Sleep(dwMilliSeconds As Long)
    Dim time As Long
    time = GetTickCount
    Do Until GetTickCount - time = dwMilliSeconds
        DoEvents
    Loop
End Sub
