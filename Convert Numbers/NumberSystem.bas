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

Private Function DecimalTo(DecNum As Double, ConvertTo As Double) As String
    Dim CurrNumber As Double, strConverted As String, Digit As Double
    strConverted = ""
    CurrNumber = DecNum
    Do Until CurrNumber = 0
        Digit = CurrNumber Mod ConvertTo
        strConverted = strConverted & BigNumToLetter(Digit)
        CurrNumber = CurrNumber \ ConvertTo
    Loop
    DecimalTo = InvertString(strConverted)
End Function

Private Function ToDecimal(Number As String, ConvertFrom As Double) As Double
    Dim i As Double, AddNum As Double, Digit As Double, NumToConvert As String
    NumToConvert = InvertString(Number)
    ToDecimal = 0
    For i = 1 To Len(Number)
        Digit = LetterToBigNum(Mid(NumToConvert, i, 1))
        AddNum = (ConvertFrom ^ (i - 1)) * Digit
        ToDecimal = ToDecimal + AddNum
    Next
End Function

Public Function ConvertNumbers(Num As String, ConvertFrom As Double, ConvertTo As Double) As String
    Dim ConvFrom As String, ConvTo As String
    ConvFrom = ToDecimal(Num, ConvertFrom)
    ConvTo = DecimalTo(CDbl(Val(ConvFrom)), ConvertTo)
    ConvertNumbers = ConvTo
End Function




Public Function GetMinNumberSystem(Number As String) As Double
    
End Function

