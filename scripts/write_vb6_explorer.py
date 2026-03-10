"""
Write frmComExplorer.frm (GBK + CRLF) and update VB6 project files.
This form uses early binding (IComPlugin) to call C# COM directly.
"""
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SLN_DIR = os.path.dirname(SCRIPT_DIR)
VB6_DIR = os.path.join(SLN_DIR, "src", "VB6ComPlugin")


def write_frm():
    """Create frmComExplorer.frm with early binding to IComPlugin"""
    content = """\
VERSION 5.00
Begin VB.Form frmComExplorer
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "C# COM Explorer (Early Binding)"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnClose
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4440
      TabIndex        =   16
      Top             =   6060
      Width           =   1335
   End
   Begin VB.CommandButton btnShutdown
      Caption         =   "Shutdown"
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton btnInitialize
      Caption         =   "Initialize"
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtResult
      Height          =   855
      Left            =   1560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   5520
      Width           =   4215
   End
   Begin VB.CommandButton btnExecute
      Caption         =   "Execute"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox txtExpression
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox txtDesc
      Height          =   855
      Left            =   1560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3360
      Width           =   4215
   End
   Begin VB.TextBox txtVersion
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2880
      Width           =   4215
   End
   Begin VB.TextBox txtName
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   4215
   End
   Begin VB.CommandButton btnRefresh
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lblTitle
      Caption         =   "C# COM Explorer"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label lblBindInfo
      Caption         =   "Binding: Early (IComPlugin via TLB)"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   720
      Width           =   5535
   End
   Begin VB.Label lblStatus
      Caption         =   "Status: Not connected"
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label lblResult
      Caption         =   "Result:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   5580
      Width           =   1215
   End
   Begin VB.Label lblExpression
      Caption         =   "Expression:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4500
      Width           =   1215
   End
   Begin VB.Label lblDesc
      Caption         =   "Description:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3420
      Width           =   1215
   End
   Begin VB.Label lblVersion
      Caption         =   "Version:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2940
      Width           =   1215
   End
   Begin VB.Label lblName
      Caption         =   "Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2460
      Width           =   1215
   End
   Begin VB.Label lblProps
      Caption         =   "--- Properties (IComPlugin) ---"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   5535
   End
   Begin VB.Label lblMethods
      Caption         =   "--- Methods ---"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   4140
      Width           =   5535
   End
   Begin VB.Line Line1
      X1              =   120
      X2              =   5880
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2
      X1              =   120
      X2              =   5880
      Y1              =   5400
      Y2              =   5400
   End
End
Attribute VB_Name = "frmComExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' Early Binding: use IComPlugin interface from CSharpComPlugin.tlb
' VB6 resolves methods via vtable at compile time (faster + IntelliSense)
' ============================================================

Private m_Plugin As CSharpComPlugin.IComPlugin

''' <summary>
''' Set the COM plugin object (early binding via IComPlugin interface)
''' </summary>
Public Sub SetPlugin(ByVal obj As Object)
    ' QueryInterface for IComPlugin - early binding starts here
    Set m_Plugin = obj

    If Not m_Plugin Is Nothing Then
        lblStatus.Caption = "Status: Connected (Early Binding)"
        lblStatus.ForeColor = &H8000&
        RefreshProperties
    Else
        lblStatus.Caption = "Status: Failed to get IComPlugin interface"
        lblStatus.ForeColor = &HFF&
        btnExecute.Enabled = False
        btnInitialize.Enabled = False
        btnShutdown.Enabled = False
    End If
End Sub

''' <summary>
''' Read all IComPlugin properties and display in text boxes
''' </summary>
Private Sub RefreshProperties()
    If m_Plugin Is Nothing Then Exit Sub

    On Error Resume Next

    ' IComPlugin.Name (DispId 1, propget)
    txtName.Text = m_Plugin.Name

    ' IComPlugin.Version (DispId 2, propget)
    txtVersion.Text = m_Plugin.Version

    ' IComPlugin.Description (DispId 3, propget)
    txtDesc.Text = m_Plugin.Description

    If Err.Number <> 0 Then
        lblStatus.Caption = "Status: Error reading properties - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub btnRefresh_Click()
    RefreshProperties
    lblStatus.Caption = "Status: Properties refreshed"
End Sub

''' <summary>
''' Call IComPlugin.Initialize() (DispId 4)
''' </summary>
Private Sub btnInitialize_Click()
    If m_Plugin Is Nothing Then Exit Sub

    On Error Resume Next
    m_Plugin.Initialize
    If Err.Number <> 0 Then
        lblStatus.Caption = "Status: Initialize failed - " & Err.Description
        Err.Clear
    Else
        lblStatus.Caption = "Status: Initialize() called OK"
    End If
    On Error GoTo 0
End Sub

''' <summary>
''' Call IComPlugin.Execute(input) (DispId 5) - returns String
''' </summary>
Private Sub btnExecute_Click()
    If m_Plugin Is Nothing Then
        txtResult.Text = "Error: Plugin not connected"
        Exit Sub
    End If

    If Len(Trim$(txtExpression.Text)) = 0 Then
        txtResult.Text = "Please enter an expression"
        Exit Sub
    End If

    On Error Resume Next
    Dim sResult As String
    sResult = m_Plugin.Execute(txtExpression.Text)
    If Err.Number <> 0 Then
        txtResult.Text = "Error: " & Err.Description
        lblStatus.Caption = "Status: Execute failed"
        Err.Clear
    Else
        txtResult.Text = sResult
        lblStatus.Caption = "Status: Execute() returned OK"
    End If
    On Error GoTo 0
End Sub

''' <summary>
''' Call IComPlugin.Shutdown() (DispId 6)
''' </summary>
Private Sub btnShutdown_Click()
    If m_Plugin Is Nothing Then Exit Sub

    On Error Resume Next
    m_Plugin.Shutdown
    If Err.Number <> 0 Then
        lblStatus.Caption = "Status: Shutdown failed - " & Err.Description
        Err.Clear
    Else
        lblStatus.Caption = "Status: Shutdown() called OK"
    End If
    On Error GoTo 0
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set m_Plugin = Nothing
End Sub
"""
    path = os.path.join(VB6_DIR, "frmComExplorer.frm")
    with open(path, "w", encoding="gbk", newline="\r\n") as f:
        f.write(content)
    print(f"[OK] {path} ({os.path.getsize(path):,} bytes)")


def update_vbp():
    """Add Form=frmComExplorer.frm and TLB reference to VBP"""
    vbp_path = os.path.join(VB6_DIR, "VB6ComPlugin.vbp")
    with open(vbp_path, "r", encoding="gbk") as f:
        lines = f.readlines()

    has_explorer_form = any("frmComExplorer" in l for l in lines)
    has_tlb_ref = any("A1B2C3D4-E5F6-7890-ABCD-EF1234567890" in l for l in lines)

    new_lines = []
    for line in lines:
        # Add TLB reference after the stdole2.tlb reference line
        if not has_tlb_ref and "stdole2.tlb" in line:
            new_lines.append(line)
            new_lines.append(
                'Reference=*\\G{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}'
                '#1.0#0#CSharpComPlugin.tlb#CSharpComPlugin 1.0 Type Library\n'
            )
            has_tlb_ref = True
            continue

        # Add form after frmCalculator line
        if not has_explorer_form and "Form=frmCalculator.frm" in line:
            new_lines.append(line)
            new_lines.append("Form=frmComExplorer.frm\n")
            has_explorer_form = True
            continue

        new_lines.append(line)

    with open(vbp_path, "w", encoding="gbk", newline="\r\n") as f:
        f.writelines(new_lines)
    print(f"[OK] {vbp_path} updated")


def update_cls():
    """Add 'explore' command to StringProcessor.cls Execute function"""
    cls_path = os.path.join(VB6_DIR, "StringProcessor.cls")
    with open(cls_path, "r", encoding="gbk") as f:
        content = f.read()

    if "frmComExplorer" in content:
        print(f"[OK] StringProcessor.cls already has explore command")
        return

    # Add explore command before the "show" case
    old_show = """\
        Case "show"
            ' Show Calculator Form with C# COM info
            If m_CalcObj Is Nothing Then
                Execute = "Error: Calculator object not set. Host must call SetCalculatorObject first."
            Else
                Dim frm As frmCalculator
                Set frm = New frmCalculator
                frm.SetInfo m_CalcName, m_CalcVersion, m_CalcDescription, m_CalcObj
                frm.Show vbModal
                Set frm = Nothing
                Execute = "Calculator form closed."
            End If"""

    new_show = """\
        Case "show"
            ' Show Calculator Form with C# COM info
            If m_CalcObj Is Nothing Then
                Execute = "Error: Calculator object not set. Host must call SetCalculatorObject first."
            Else
                Dim frm As frmCalculator
                Set frm = New frmCalculator
                frm.SetInfo m_CalcName, m_CalcVersion, m_CalcDescription, m_CalcObj
                frm.Show vbModal
                Set frm = Nothing
                Execute = "Calculator form closed."
            End If

        Case "explore"
            ' Open COM Explorer Form (early binding via IComPlugin TLB)
            If m_CalcObj Is Nothing Then
                Execute = "Error: Calculator object not set. Host must call SetCalculatorObject first."
            Else
                Dim frmEx As frmComExplorer
                Set frmEx = New frmComExplorer
                frmEx.SetPlugin m_CalcObj
                frmEx.Show vbModal
                Set frmEx = Nothing
                Execute = "COM Explorer form closed."
            End If"""

    content = content.replace(old_show, new_show)

    # Also update the help text in Description and error messages
    content = content.replace(
        "Commands: upper/lower/reverse/len/words/calc/show",
        "Commands: upper/lower/reverse/len/words/calc/show/explore"
    )
    content = content.replace(
        "Commands: upper, lower, reverse, len, words, calc, show",
        "Commands: upper, lower, reverse, len, words, calc, show, explore"
    )

    with open(cls_path, "w", encoding="gbk", newline="\r\n") as f:
        f.write(content)
    print(f"[OK] {cls_path} updated with 'explore' command")


if __name__ == "__main__":
    write_frm()
    update_vbp()
    update_cls()
    print("\nAll VB6 files updated!")
