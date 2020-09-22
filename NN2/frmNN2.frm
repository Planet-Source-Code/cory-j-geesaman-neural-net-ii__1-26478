VERSION 5.00
Begin VB.Form frmNN2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NN2 - By: Cory J. Geesaman"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNN2.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Height          =   1020
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Text            =   "Input"
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Output:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Input:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Time:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   390
   End
End
Attribute VB_Name = "frmNN2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Engrams(0 To 256) As clsEngram, ExitLoop As Boolean

Private Sub Command1_Click()
Dim i, j, e, e2, tC
If Command1.Caption = "&Start" Then
Command1.Caption = "&Stop"
ExitLoop = False
tC = GetTickCount
Text2.Text = ""
Text3.Text = 0
e = LBound(Engrams)
i = 1
Do
StartOff:
DoEvents
e2 = Engrams(e).GetStrongestBond(e)
DoEvents
If Mid(Text1.Text, i, 1) = Engrams(e2).Data Then
Engrams(e).ChangeBondStrength Engrams(e2), Engrams(e).GetBondStrength(Engrams(e2)) + 1
Text2.Text = Text2.Text & Engrams(e2).Data
e = e2
i = i + 1
Else
Engrams(e).ChangeBondStrength Engrams(e2), Engrams(e).GetBondStrength(Engrams(e2)) - 1
GoTo StartOff
End If
DoEvents
Loop Until i > Len(Text1.Text) Or ExitLoop = True
If ExitLoop = True Then
Text2.Text = ""
Exit Sub
End If
Text3.Text = GetTickCount - tC
Command1.Caption = "&Start"
Else
Command1.Caption = "&Start"
ExitLoop = True
End If
End Sub

Private Sub Form_Load()
Dim i, j
i = LBound(Engrams)
Do
Set Engrams(i) = New clsEngram
If i <> LBound(Engrams) Then Engrams(i).Data = Chr(i - 1)
i = i + 1
Loop Until i > UBound(Engrams)
i = LBound(Engrams)
Do
j = LBound(Engrams) + 1
Do
If j <> i Then Engrams(i).AddConnection Engrams(j), 0
j = j + 1
Loop Until j > UBound(Engrams)
i = i + 1
Loop Until i > UBound(Engrams)
End Sub
