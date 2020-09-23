Attribute VB_Name = "mdlMain"
Option Explicit

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200

Public Sub Main()

    On Error Resume Next
    
    Dim iccex As tagInitCommonControlsEx
    
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    
    InitCommonControlsEx iccex
    
    On Error GoTo 0
    Form1.Show
    
End Sub


