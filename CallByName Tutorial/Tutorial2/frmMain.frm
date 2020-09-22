VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Tutorial 2 - Properties"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   398
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   105
      Top             =   1095
   End
   Begin VB.CommandButton cmdEnableTimer 
      Caption         =   "Enable Timer"
      Height          =   390
      Left            =   90
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdChangeCaption 
      Caption         =   "Change form caption"
      Height          =   390
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChangeCaption_Click()
    CallByName frmMain, "Caption", VbLet, "CallByName - Tutorial 2"
End Sub

Private Sub cmdEnableTimer_Click()
    CallByName tmrMove, "Enabled", VbLet, Not CallByName(tmrMove, "Enabled", VbGet)
End Sub

Private Sub tmrMove_Timer()
    CallByName cmdEnableTimer, "Left", VbLet, CInt(Rnd(frmMain.ScaleWidth)) * 100
    CallByName cmdEnableTimer, "Top", VbLet, CInt(Rnd(frmMain.ScaleHeight)) * 100
End Sub
