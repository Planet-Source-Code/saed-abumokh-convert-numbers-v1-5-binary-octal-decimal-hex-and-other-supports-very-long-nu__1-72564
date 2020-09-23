VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Numbers v1.5 (supports very long numbers)"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSwapNumbers 
      Caption         =   "&Swap Numbers"
      Height          =   285
      Left            =   4980
      TabIndex        =   19
      Top             =   870
      Width           =   1365
   End
   Begin VB.CommandButton cmdSwap 
      Caption         =   "S&wap"
      Height          =   285
      Left            =   2783
      TabIndex        =   18
      Top             =   3000
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Binary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   563
      TabIndex        =   16
      Top             =   1800
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Octal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1943
      TabIndex        =   15
      Top             =   1800
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Decimal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3293
      TabIndex        =   14
      Top             =   1800
      Width           =   1275
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Hexadecimal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4643
      TabIndex        =   13
      Top             =   1800
      Width           =   1275
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Hexadecimal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4643
      TabIndex        =   12
      Top             =   3990
      Width           =   1275
   End
   Begin VB.CommandButton Command6 
      Caption         =   "D&ecimal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3293
      TabIndex        =   11
      Top             =   3990
      Width           =   1275
   End
   Begin VB.CommandButton Command7 
      Caption         =   "O&ctal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1943
      TabIndex        =   10
      Top             =   3990
      Width           =   1275
   End
   Begin VB.CommandButton Command8 
      Caption         =   "B&inary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   563
      TabIndex        =   9
      Top             =   3990
      Width           =   1275
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5790
      Width           =   6225
   End
   Begin VB.TextBox txtTo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4485
      TabIndex        =   6
      Text            =   "30"
      Top             =   4470
      Width           =   615
   End
   Begin VB.TextBox txtFrom 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4485
      TabIndex        =   2
      Text            =   "16"
      Top             =   2250
      Width           =   615
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Co&nvert"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2528
      TabIndex        =   1
      Top             =   5340
      Width           =   1425
   End
   Begin VB.TextBox txtConvert 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "CD333B35AA395EEE22DDDF0F3CDD92E519492BD7D72A615788543A"
      Top             =   390
      Width           =   6225
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Or choose any number between 2 and 36:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1155
      TabIndex        =   20
      Top             =   4530
      Width           =   3075
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   5520
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Convert To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2790
      TabIndex        =   7
      Top             =   3660
      Width           =   870
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Or choose any number between 2 and 36:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1155
      TabIndex        =   5
      Top             =   2340
      Width           =   3045
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type any number here:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Convert From:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2715
      TabIndex        =   3
      Top             =   1500
      Width           =   1050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Long

Private Sub cmdConvert_Click()
    If txtFrom.Text = "" Then
        FlashWindow txtFrom, RGB(255, 64, 64), 3, 250, 250, 600
        txtFrom.SetFocus
    ElseIf IsNumeric(txtFrom) And (txtFrom.Text > 36) Or (txtFrom.Text < 2) Then
        Label4.FontBold = True
        FlashWindow txtFrom, RGB(255, 64, 64), 3, 250, 250, 600
        Label4.FontBold = False
        txtFrom.SetFocus
    ElseIf txtTo.Text = "" Then
        FlashWindow txtTo, RGB(255, 64, 64), 3, 250, 250, 600
        txtTo.SetFocus
    ElseIf IsNumeric(txtTo) And (txtTo.Text > 36) Or (txtTo.Text < 2) Then
        Label5.FontBold = True
        FlashWindow txtTo, RGB(255, 64, 64), 3, 250, 250, 600
        Label5.FontBold = False
        txtFrom.SetFocus
    Else
        If GetMaxNumberSystem(txtConvert.Text) > CDbl(Val(txtFrom.Text)) Then _
            txtFrom.Text = GetMaxNumberSystem(txtConvert.Text)
        txtConvert.Text = Replace(txtConvert.Text, " ", "0")
        txtResult.Text = ConvertNumbers(txtConvert.Text, txtFrom.Text, txtTo.Text)
        
    End If
End Sub

Private Sub cmdSwap_Click()
    Dim textFrom As Integer
    textFrom = Val(txtFrom.Text)
    txtFrom.Text = txtTo.Text
    txtTo.Text = textFrom
End Sub

Private Sub cmdSwapNumbers_Click()
    Dim textConvert As String
    textConvert = txtConvert.Text
    txtConvert.Text = txtResult.Text
    txtResult.Text = textConvert
End Sub

Private Sub Command1_Click()
    txtFrom.Text = 2
End Sub

Private Sub Command2_Click()
    txtFrom.Text = 8
End Sub

Private Sub Command3_Click()
    txtFrom.Text = 10
End Sub

Private Sub Command4_Click()
    txtFrom.Text = 16
End Sub

Private Sub Command5_Click()
    txtTo.Text = 16
End Sub

Private Sub Command6_Click()
    txtTo.Text = 10
End Sub

Private Sub Command7_Click()
    txtTo.Text = 8
End Sub

Private Sub Command8_Click()
    txtTo.Text = 2
End Sub

Private Sub Form_Load()
    i = 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub txtConvert_Change()
    Dim GetNumberSystem As Double
    
    If GetMaxNumberSystem(txtConvert.Text) > CDbl(Val(txtFrom.Text)) Then
        GetNumberSystem = GetMaxNumberSystem(txtConvert.Text)
    Else
        GetNumberSystem = txtFrom.Text
    End If
    
    Dim CheckErr As String
End Sub

Private Sub txtConvert_KeyPress(KeyAscii As Integer)
    If ((KeyAscii <> 8) And (KeyAscii <> 32)) And (((KeyAscii < vbKey0) Or (KeyAscii > vbKey9)) And ((KeyAscii < 65) Or (KeyAscii > 90)) And ((KeyAscii < 97) Or (KeyAscii > 122))) Then
        KeyAscii = 0
    End If
    txtConvert_Change
End Sub

Private Sub txtFrom_Change()
    txtConvert_Change
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) And ((KeyAscii < vbKey0) Or (KeyAscii > vbKey9)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTo_Change()
    txtConvert_Change
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) And ((KeyAscii < vbKey0) Or (KeyAscii > vbKey9)) Then
        KeyAscii = 0
    End If
End Sub
