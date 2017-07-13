VERSION 5.00
Begin VB.Form frm1 
   BackColor       =   &H8000000A&
   Caption         =   "Calculator"
   ClientHeight    =   5850
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4065
   ForeColor       =   &H8000000A&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ComFact 
      Caption         =   "n!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   30
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton ComRcpl 
      Caption         =   "1/n"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   29
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton ComSqrt 
      Caption         =   "Sqrt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   28
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton ComLog 
      Caption         =   "Ln"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   27
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton ComExp 
      Caption         =   "Exp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   26
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton ComSin 
      Caption         =   "Sin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   25
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton ComMod 
      Caption         =   "Mod"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   24
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton ComCos 
      Caption         =   "Cos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   23
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Compercent 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   22
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton ComTan 
      Caption         =   "Tan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   21
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Comans 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   17
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton ComCE 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Comdiv 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   15
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Com0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   14
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Compt 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   13
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Com9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Commult 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   11
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton ComC 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Comsub 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   9
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Com7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Com8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Com4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Com5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Com6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Comadd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Com2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Com3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   1
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Com1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lbl2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblans 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1080
      TabIndex        =   20
      Top             =   720
      Width           =   2760
   End
   Begin VB.Label display 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   3735
   End
   Begin VB.Menu View 
      Caption         =   "View"
      NegotiatePosition=   1  'Left
      Begin VB.Menu basic 
         Caption         =   "Standard"
         Shortcut        =   ^B
      End
      Begin VB.Menu Scientific 
         Caption         =   "Scientific"
         Shortcut        =   ^S
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'for basic
Dim fnum As Double
Dim lnum As Double
Dim operation As String





Private Sub Command11_Click()

End Sub

Private Sub Command12_Click()

End Sub


Private Sub About_Click()
MsgBox "Scientific Calculator V1.0 Designed by SUSANTA"
End Sub

Private Sub basic_Click()
frm1.Height = 6735
frm1.Width = 4305
lblans.Width = 2760
display.Width = 3735
basic.Enabled = False
Scientific.Enabled = True
lblans.Caption = "0"
End Sub

Private Sub Com0_Click()
 'when 0 is displayed in display, you cannot input 0 again
If lblans.Caption = "0" Then
lblans.Caption = lblans.Caption & 0
lblans.Caption = 0 + lblans.Caption
'if DIsplay is not equal to 0, then u can input 0 next to 0(only after "." and "0.0# etc")
ElseIf lblans.Caption <> "0" Then
lblans.Caption = lblans.Caption & 0
End If
End Sub



Private Sub Com1_Click()
 lblans.Caption = lblans.Caption & 1
lblans.Caption = 0 + lblans.Caption
    
End Sub


Private Sub Com2_Click()
  lblans.Caption = lblans.Caption & 2
lblans.Caption = 0 + lblans.Caption
End Sub

Private Sub Com3_Click()
  lblans.Caption = lblans.Caption & 3
lblans.Caption = 0 + lblans.Caption
End Sub

Private Sub Com4_Click()
   lblans.Caption = lblans.Caption & 4
lblans.Caption = 0 + lblans.Caption
End Sub

Private Sub Com5_Click()
 lblans.Caption = lblans.Caption & 5
lblans.Caption = 0 + lblans.Caption
End Sub

Private Sub Com6_Click()
  lblans.Caption = lblans.Caption & 6
lblans.Caption = 0 + lblans.Caption
End Sub

Private Sub Com7_Click()
 lblans.Caption = lblans.Caption & 7
lblans.Caption = 0 + lblans.Caption
End Sub

Private Sub Com8_Click()
 lblans.Caption = lblans.Caption & 8
lblans.Caption = 0 + lblans.Caption
End Sub

Private Sub Com9_Click()
 lblans.Caption = lblans.Caption & 9
lblans.Caption = 0 + lblans.Caption
End Sub

Private Sub Comadd_Click()
    'addition
Compt.Enabled = True
fnum = lblans.Caption
lbl2.Caption = lblans.Caption
lblans.Caption = "0"
operation = "+"
lbl2.Caption = lbl2.Caption + " " + operation
End Sub

Private Sub Comans_Click()
Compt.Enabled = True
lnum = lblans.Caption

'----------------
'basic operations
'----------------
If operation = "+" Then
lblans.Caption = fnum + lnum
lbl2.Caption = ""

ElseIf operation = "-" Then
 lblans.Caption = fnum - lnum
lbl2.Caption = ""

ElseIf operation = "x" Then
lblans.Caption = fnum * lnum
lbl2.Caption = ""

ElseIf operation = "/" Then
lblans.Caption = fnum / lnum
lbl2.Caption = ""
ElseIf operation = "^" Then
lblans.Caption = fnum ^ lnum
lbl2.Caption = ""
ElseIf operation = "Mod" Then
lblans.Caption = fnum Mod lnum
lbl2.Caption = ""
ElseIf operation = "Sin" Then
lblans.Caption = Math.Sin(lblans.Caption)
lbl2.Caption = ""
ElseIf operation = "Cos" Then
lblans.Caption = Math.Cos(lblans.Caption)
lbl2.Caption = ""
ElseIf operation = "Tan" Then
lblans.Caption = Math.Tan(lblans.Caption)
lbl2.Caption = ""
ElseIf operation = "Log" Then
lblans.Caption = Math.Log(lblans.Caption)
lbl2.Caption = ""
ElseIf operation = "Sqrt" Then
lblans.Caption = Math.Sqr(lblans.Caption)
lbl2.Caption = ""
ElseIf operation = "1/n" Then
lblans.Caption = 1 / fnum
lbl2.Caption = ""
End If
End Sub

Private Sub ComBin_Click()


End Sub

Private Sub ComC_Click()
'command cancel
lblans.Caption = Mid(lblans.Caption, 1, Len(lblans.Caption) - 1)
If lblans.Caption = "" Then
lblans.Caption = "0"
'decimal backspace
ElseIf lblans.Caption = Format(lblans.Caption, "#") Then
Compt.Enabled = True
ElseIf lblans.Caption = "0" Then
Compt.Enabled = True
End If
End Sub

Private Sub ComCE_Click()
'clear
Compt.Enabled = True
lblans.Caption = "0"
    
End Sub

Private Sub ComCos_Click()
Compt.Enabled = True
fnum = lblans.Caption
lbl2.Caption = lblans.Caption
lblans.Caption = "0"
operation = "Cos"
lbl2.Caption = lbl2.Caption + " " + operation
End Sub

Private Sub Comdiv_Click()
 Compt.Enabled = True
fnum = lblans.Caption
lbl2.Caption = lblans.Caption
lblans.Caption = "0"
operation = "/"
lbl2.Caption = lbl2.Caption + " " + operation
End Sub

Private Sub ComExp_Click()
Compt.Enabled = True
fnum = lblans.Caption
lbl2.Caption = lblans.Caption
lblans.Caption = "0"
operation = "^"
lbl2.Caption = lbl2.Caption + " " + operation
End Sub

Private Sub ComFact_Click()
Dim A As Double
Dim B As Double
Dim M As Double
Dim Fac As Double

If lblans.Caption = "Syntax Error" Then
lblans.Caption = "Syntax Error"
Beep
Else
A = Val(lblans.Caption)
M = 1
Fac = 1
Do While M <= A
Fac = Fac * M
M = M + 1
Loop
lblans.Caption = Format(Fac, "#")
End If
End Sub

Private Sub ComLog_Click()
Compt.Enabled = True
fnum = lblans.Caption
lbl2.Caption = lblans.Caption
lblans.Caption = "0"
operation = "Log"
lbl2.Caption = lbl2.Caption + " " + operation
End Sub

Private Sub ComMod_Click()
Compt.Enabled = True
fnum = lblans.Caption
lbl2.Caption = lblans.Caption
lblans.Caption = "0"
operation = "Mod"
lbl2.Caption = lbl2.Caption + " " + operation
End Sub

Private Sub Commult_Click()
 Compt.Enabled = True
fnum = lblans.Caption
lbl2.Caption = lblans.Caption
lblans.Caption = "0"
operation = "x"
lbl2.Caption = lbl2.Caption + " " + operation
End Sub

Private Sub Compercent_Click()
Compt.Enabled = True
fnum = lblans.Caption
lbl2.Caption = lblans.Caption
lblans.Caption = fnum / 100
operation = "%"
lbl2.Caption = lbl2.Caption + " " + operation
End Sub

Private Sub Compt_Click()
'decimal point
lblans.Caption = lblans.Caption & "."
Compt.Enabled = False
If lblans.Caption = "." Then
lblans.Caption = Format(lblans.Caption, "0.#")
ElseIf lblans.Caption = operation Then
Compt.Enabled = True


End If
End Sub

Private Sub ComRcpl_Click()
Compt.Enabled = True
fnum = lblans.Caption
lblans.Caption = 1 / fnum
operation = "1/n"

End Sub

Private Sub ComSin_Click()
Compt.Enabled = True
fnum = lblans.Caption
lbl2.Caption = lblans.Caption
lblans.Caption = "0"
operation = "Sin"
lbl2.Caption = lbl2.Caption + " " + operation
End Sub

Private Sub ComSqrt_Click()
Compt.Enabled = True
fnum = lblans.Caption
lbl2.Caption = lblans.Caption
lblans.Caption = "0"
operation = "Sqrt"
lbl2.Caption = lbl2.Caption + " " + operation
End Sub

Private Sub Comsub_Click()
Compt.Enabled = True
fnum = lblans.Caption
lbl2.Caption = lblans.Caption
lblans.Caption = "0"
operation = "-"
lbl2.Caption = lbl2.Caption + " " + operation
End Sub

Private Sub ComTan_Click()
Compt.Enabled = True
fnum = lblans.Caption
lbl2.Caption = lblans.Caption
lblans.Caption = "0"
operation = "Tan"
lbl2.Caption = lbl2.Caption + " " + operation
End Sub

Private Sub Exit_Click()
End

End Sub

Private Sub Form_Load()
basic.Enabled = False
lblans.Enabled = False
lblans.Caption = "0"
   
End Sub


Private Sub Scientific_Click()


frm1.Height = 6735
frm1.Width = 6270
Comans.Width = 1815
display.Width = 5775
lblans.Width = 4695
Scientific.Enabled = False
basic.Enabled = True
lblans.Caption = "0"
End Sub

Private Sub Text1_Change()

End Sub
