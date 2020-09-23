Attribute VB_Name = "NumberSystem"
Option Explicit

Private Function InvertString(str As String) As String
    Dim i As Double
    For i = Len(str) To 1 Step -1
        InvertString = InvertString & Mid(str, i, 1)
    Next
End Function

Private Function BigNumToLetter(ByVal Num As Double) As String
    If Num < 10 Then
        BigNumToLetter = CStr(Num)
    ElseIf (Num >= 10) And (Num <= 35) Then
        BigNumToLetter = Chr$(Num - 10 + 65)
    End If
End Function

Private Function LetterToBigNum(ByVal sLetter As String) As Double
    Dim Letter As String
    Letter = UCase(sLetter)
    If (Asc(Letter) >= vbKey0) And (Asc(Letter) <= vbKey9) Then
        LetterToBigNum = Val(Letter)
    ElseIf (Asc(Letter) >= vbKeyA) And (Asc(Letter) <= vbKeyZ) Then
        LetterToBigNum = Asc(Letter) - 65 + 10
    End If
End Function

Private Function DecimalTo(DecNum As String, ConvertTo As String) As String
    Dim CurrNumber As String, strConverted As String, Digit As String
    strConverted = ""
    CurrNumber = DecNum
    Do Until CurrNumber = 0
        Digit = Modulo((CurrNumber), (ConvertTo))
        strConverted = strConverted & BigNumToLetter(Digit)
        CurrNumber = DivideNonRestoring((CurrNumber), (ConvertTo))
    Loop
    DecimalTo = InvertString(strConverted)
End Function

Private Function ToDecimal(Number As String, ConvertFrom As String) As String
    Dim i As Integer, AddNum As String, Digit As String, NumToConvert As String
    NumToConvert = InvertString(Number)
    ToDecimal = 0
    For i = 1 To Len(Number)
        Digit = LetterToBigNum(Mid(NumToConvert, i, 1))
        AddNum = Multiply(Power((ConvertFrom), CStr((i - 1))), Digit)
        ToDecimal = Add(ToDecimal, AddNum)
    Next
End Function

Public Function ConvertNumbers(Num As String, ConvertFrom As String, ConvertTo As String) As String
    Dim ConvFrom As String, ConvTo As String
    ConvFrom = ToDecimal(Num, ConvertFrom)
    ConvTo = DecimalTo(ConvFrom, ConvertTo)
    ConvertNumbers = ConvTo
End Function


Public Function GetMaxNumberSystem(Number As String) As Double
    Dim i As Integer
    Dim Digits() As Double
    If Number = "" Then Number = "0"
    ReDim Digits(1 To Len(Number)) As Double
    For i = 1 To Len(Number)
        Digits(i) = LetterToBigNum(Mid(Number, i, 1))
    Next
    GetMaxNumberSystem = Maximum(Digits) + 1
End Function

Private Function Maximum(NumArray() As Double) As Double
    If UBound(NumArray) - LBound(NumArray) = 0 Then
        Maximum = 0
    ElseIf UBound(NumArray) - LBound(NumArray) = 1 Then
        Maximum = LBound(NumArray)
    ElseIf UBound(NumArray) - LBound(NumArray) > 1 Then
        Dim i As Integer, CurrNumber As Double, CompareNumber As Double
        CurrNumber = CompareMax(NumArray(LBound(NumArray)), NumArray(LBound(NumArray) + 1))
        For i = LBound(NumArray) To UBound(NumArray) - 1
            CompareNumber = CompareMax(CurrNumber, NumArray(i + 1))
            CurrNumber = CompareMax(NumArray(i), CompareNumber)
        Next
        Maximum = CompareNumber
    End If
End Function

Private Function CompareMax(Num1 As Double, Num2 As Double) As Double
    '(|x+y|)/2 + (|x-y|)/2
    CompareMax = (Num1 + Num2) / 2 + (Abs(Num1 - Num2)) / 2
End Function
