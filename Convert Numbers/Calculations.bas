Attribute VB_Name = "Calculations"
Option Explicit

Private Function Max(ByVal Num1 As Double, ByVal Num2 As Double) As Double
    Max = (Num1 + Num2) / 2 + (Abs(Num1 - Num2)) / 2
End Function

Private Function FillWithZeroes(Number As String, Length As Long) As String
    On Error Resume Next
    FillWithZeroes = String(Length - Len(Number), "0") & Number
End Function

Private Function LTrimChar(Number As String, Char As String, Optional ByVal AssumeAsNoneAsChar As Boolean = True) As String
    Dim i As Integer
    For i = 1 To Len(Number)
        If Mid(Number, i, 1) <> Char Then Exit For
    Next
    i = i - 1
    LTrimChar = Right(Number, Len(Number) - i)
    If AssumeAsNoneAsChar = True Then If LTrimChar = "" Then LTrimChar = Char
End Function

Private Function MultiplyBy1Digit(Num1 As String, Digit As String) As String ' a * b and b < 10
    Dim i As Integer
    Dim Result As String
    Dim Counter As String
    For i = 1 To (Digit)
        Result = Add(Result, Num1, 2)
    Next
    MultiplyBy1Digit = Result
End Function

Public Function Multiply(Num1 As String, Num2 As String) As String
    Dim Numbers() As String
    Dim i As Integer
    Dim Result As String
    Dim Num2Inverted As String
    
    
    ReDim Numbers(1 To Len(Num2)) As String
    
    Num2Inverted = InvertString(Num2)
    For i = 1 To Len(Num2Inverted)
        Numbers(i) = MultiplyBy1Digit(Num1, (Mid(Num2Inverted, i, 1))) & String(i - 1, "0")
    Next
    
    Result = Numbers(1)
    
    If UBound(Numbers) > 1 Then
        For i = 2 To UBound(Numbers)
            Result = Add(Result, Numbers(i), 10)
        Next
    End If
    Multiply = Result
End Function

Public Function Power(Num1 As String, Num2 As String) As String ' a ^ b
    Dim i As Double
    Dim Result As String
    Dim Counter As String
    Result = "1"
    For i = 1 To Val(Num2)
        Result = Multiply(Result, Num1)
    Next
    Power = Result
End Function
Private Function InvertString(str As String) As String
    Dim i As Double
    For i = Len(str) To 1 Step -1
        InvertString = InvertString & Mid(str, i, 1)
    Next
End Function

Private Function IsValueZero(Number As String) As Boolean
    Dim i As Integer
    IsValueZero = True
    For i = 1 To Len(Number)
        If Mid(Number, i, 1) <> "0" Then
            IsValueZero = False
        End If
    Next
End Function

'    Dim TwoDigits As String
'    Dim i As Integer
'    Dim DigitResult As String
'    Dim Result As String
'    Dim DigitRest As String
'
'    TwoDigits = "0" & Mid(Numerator, 1, 1)
'
'    For i = 1 To Len(Numerator)
'        DigitResult = TwoDigits \ Denominator
'        DigitRest = TwoDigits Mod Denominator
'        Result = Result & DigitResult
'        TwoDigits = DigitRest & Mid(Numerator, i + 1, 1)
'    Next
'    DivideNonRestoring = LTrimChar(Result, "0")


'summary:
Public Function DivideNonRestoring(Numerator As String, Denominator As String) As String ' a \ b --> int(a / b)
    Dim TwoDigits As String
    Dim i As Integer
    Dim Result As String
    
    TwoDigits = "0" & Mid(Numerator, 1, 1)
    
    For i = 1 To Len(Numerator)
        Result = Result & (TwoDigits \ Denominator)
        TwoDigits = (TwoDigits Mod Denominator) & Mid(Numerator, i + 1, 1)
    Next
    DivideNonRestoring = LTrimChar(Result, "0")
End Function


'    Dim TwoDigits As String
'    Dim i As Integer
'    Dim DigitResult As String
'    Dim Result As String
'    Dim DigitRest As String
'
'    TwoDigits = "0" & Mid(Numerator, 1, 1)
'
'    For i = 1 To Len(Numerator)
'        DigitResult = TwoDigits \ Denominator
'        DigitRest = TwoDigits Mod Denominator
'        Result = Result & DigitResult
'        TwoDigits = DigitRest & Mid(Numerator, i + 1, 1)
'    Next
'    Modulo = LTrimChar(DigitRest, "0")

'summary:
Public Function Modulo(Numerator As String, Denominator As String) As String ' a mod b
    Dim TwoDigits As String
    Dim i As Integer
    
    TwoDigits = "0" & Mid(Numerator, 1, 1)
    
    For i = 1 To Len(Numerator)
        TwoDigits = (TwoDigits Mod Denominator) & Mid(Numerator, i + 1, 1)
    Next
    Modulo = LTrimChar((TwoDigits Mod Denominator), "0")
End Function

Public Function Add(ByVal Num1 As String, ByVal Num2 As String, Optional ByVal GroupNumsLength As Byte = 15) As String ' a + b
    ' a + b
    Dim Group1 As Double, Group2 As Double, GroupResult As Double
    Dim NumZeroes As Integer
    Dim Number1 As String, Number2 As String
    Dim Result As String, CarryNumber As String
    Dim i As Integer, MaxString As Long
        
    NumZeroes = GroupNumsLength - 1
    
    Number1 = FillWithZeroes(Num1, Celling(Len(Num1), GroupNumsLength))
    Number2 = FillWithZeroes(Num2, Celling(Len(Num2), GroupNumsLength))

    MaxString = Max(Len(Number1), Len(Number2))
    Number1 = FillWithZeroes(Number1, MaxString)
    Number2 = FillWithZeroes(Number2, MaxString)
        
    For i = 1 To MaxString Step GroupNumsLength
        Group1 = Val(Mid(Number1, i, GroupNumsLength))
        Group2 = Val(Mid(Number2, i, GroupNumsLength))
        GroupResult = Group1 + Group2
        If GroupResult < 10 ^ GroupNumsLength Then
            Result = Result & Format(GroupResult, String(GroupNumsLength, "0"))
            CarryNumber = CarryNumber & "0" & String(NumZeroes, "0")
        ElseIf GroupResult >= 10 ^ GroupNumsLength Then
            Result = Result & Format(GroupResult - (10 ^ GroupNumsLength), String(GroupNumsLength, "0"))
            CarryNumber = CarryNumber & "1" & String(NumZeroes, "0")
        End If
    Next
    
    Do Until IsValueZero(CarryNumber)
    
        CarryNumber = CarryNumber & "0"
        Number1 = Result: Result = ""
        Number2 = CarryNumber: CarryNumber = ""
        
        Number1 = FillWithZeroes(Number1, Celling(Len(Number1), GroupNumsLength))
        Number2 = FillWithZeroes(Number2, Celling(Len(Number2), GroupNumsLength))
    
        MaxString = Max(Len(Number1), Len(Number2))
        Number1 = FillWithZeroes(Number1, MaxString)
        Number2 = FillWithZeroes(Number2, MaxString)
        
        For i = 1 To MaxString Step GroupNumsLength
            Group1 = Val(Mid(Number1, i, GroupNumsLength))
            Group2 = Val(Mid(Number2, i, GroupNumsLength))
            GroupResult = Group1 + Group2
            If GroupResult < 10 ^ GroupNumsLength Then
                Result = Result & Format(GroupResult, String(GroupNumsLength, "0"))
                CarryNumber = CarryNumber & "0" & String(NumZeroes, "0")
            ElseIf GroupResult >= 10 ^ GroupNumsLength Then
                Result = Result & Format(GroupResult - (10 ^ GroupNumsLength), String(GroupNumsLength, "0"))
                CarryNumber = CarryNumber & "1" & String(NumZeroes, "0")
            End If
        Next
    Loop
    
    Add = LTrimChar(Result, "0")
End Function


Private Function Celling(ByVal Number As Double, ByVal Steps As Integer) As Double
    Celling = (Int(Number / Steps) * Steps) + (Steps * (Sgn(Number Mod Steps)))
End Function

