VERSION 5.00
Begin VB.Form frmCalculator 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "C# COM Component Info"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton btnCalc 
      Caption         =   "Calculate"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtResult 
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4200
      Width           =   3735
   End
   Begin VB.TextBox txtExpression 
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox txtDesc 
      Height          =   855
      Left            =   1440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2640
      Width           =   3735
   End
   Begin VB.TextBox txtVersion 
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label lblTitle 
      Caption         =   "C# COM Component"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label lblStatus 
      Caption         =   "Ready"
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label lblResult 
      Caption         =   "Result:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4260
      Width           =   1095
   End
   Begin VB.Label lblExpression 
      Caption         =   "Expression:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3780
      Width           =   1095
   End
   Begin VB.Label lblDesc 
      Caption         =   "Description:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2220
      Width           =   1095
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5280
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5280
      Y1              =   3480
      Y2              =   3480
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CalcObj As Object

Public Sub SetInfo(ByVal sName As String, ByVal sVersion As String, ByVal sDesc As String, ByVal objCalc As Object)
    txtName.Text = sName
    txtVersion.Text = sVersion
    txtDesc.Text = sDesc
    Set m_CalcObj = objCalc
    
    If Not m_CalcObj Is Nothing Then
        lblStatus.Caption = "Connected: " & sName & " v" & sVersion
    Else
        lblStatus.Caption = "Calculator not available"
        btnCalc.Enabled = False
    End If
End Sub

Private Sub btnCalc_Click()
    If m_CalcObj Is Nothing Then
        txtResult.Text = "Error: Calculator not connected"
        Exit Sub
    End If
    
    If Len(Trim$(txtExpression.Text)) = 0 Then
        txtResult.Text = "Please enter an expression"
        Exit Sub
    End If
    
    On Error Resume Next
    Dim sResult As String
    sResult = m_CalcObj.Execute(txtExpression.Text)
    If Err.Number <> 0 Then
        txtResult.Text = "Error: " & Err.Description
        Err.Clear
    Else
        txtResult.Text = sResult
    End If
    On Error GoTo 0
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set m_CalcObj = Nothing
End Sub
