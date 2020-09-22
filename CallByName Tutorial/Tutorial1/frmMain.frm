VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "CallByName - Tutorial 1"
   ClientHeight    =   915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   915
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Calculate"
      Default         =   -1  'True
      Height          =   390
      Left            =   1080
      TabIndex        =   4
      Top             =   495
      Width           =   1695
   End
   Begin VB.ComboBox cmbAction 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   135
      Width           =   1695
   End
   Begin VB.TextBox txtValue2 
      Height          =   300
      Left            =   2865
      TabIndex        =   1
      Top             =   135
      Width           =   1035
   End
   Begin VB.TextBox txtValue1 
      Height          =   285
      Left            =   75
      TabIndex        =   0
      Top             =   135
      Width           =   930
   End
   Begin VB.Label lblResult 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3960
      TabIndex        =   3
      Top             =   135
      Width           =   2640
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExecute_Click()
    lblResult.Caption = "= " & CallByName(frmMain, cmbAction.Text, VbMethod, txtValue1, txtValue2)
End Sub

Private Sub Form_Load()
    With cmbAction
        .AddItem "Multiply"
        .AddItem "Minus"
        .AddItem "DivideBy"
        .AddItem "Plus"
        .ListIndex = 0
    End With
End Sub

Public Function Multiply(lngValue1 As Long, lngValue2 As Long) As Long
    Multiply = lngValue1 * lngValue2
End Function

Public Function Minus(lngValue1 As Long, lngValue2 As Long) As Long
    Minus = lngValue1 - lngValue2
End Function

Public Function DivideBy(lngValue1 As Long, lngValue2 As Long) As Long
    DivideBy = lngValue1 / lngValue2
End Function

Public Function Plus(lngValue1 As Long, lngValue2 As Long) As Long
    Plus = lngValue1 + lngValue2
End Function
