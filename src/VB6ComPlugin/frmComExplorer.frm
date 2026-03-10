VERSION 5.00
Begin VB.Form frmComExplorer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "C# COM Explorer (Reg-Free COM)"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton btnClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4440
      TabIndex        =   18
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton btnCreateViaManifest 
      Caption         =   "Create via Manifest"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton btnShutdown 
      Caption         =   "Shutdown"
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton btnInitialize 
      Caption         =   "Initialize"
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox txtResult 
      Height          =   855
      Left            =   1560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   5640
      Width           =   4215
   End
   Begin VB.CommandButton btnExecute 
      Caption         =   "Execute"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox txtExpression 
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Top             =   5160
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
      Top             =   4560
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
      TabIndex        =   20
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label lblBindInfo 
      Caption         =   "Late Binding / Reg-Free COM (No Registration)"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   720
      Width           =   5535
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status: Not connected"
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label lblResult 
      Caption         =   "Result:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   5700
      Width           =   1215
   End
   Begin VB.Label lblExpression 
      Caption         =   "Expression:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   5220
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
      Top             =   4620
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
      Y1              =   6480
      Y2              =   6480
   End
End
Attribute VB_Name = "frmComExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' Late Binding + Reg-Free COM
' No TLB registration, no COM registration needed
' VB6ComPlugin.manifest declares <dependency> on CSharpComPlugin
' SxS resolves CreateObject("ComPluginDemo.Calculator") via manifest
'
' VB6 IDE debug mode:
'   Uses modActCtx to manually activate CSharpComPlugin.manifest
'   via Win32 Activation Context API (CreateActCtx/ActivateActCtx)
' ============================================================

Private m_Plugin As Object

''' <summary>
''' Receive COM object from C# host (passed via SetCalculatorObject)
''' </summary>
Public Sub SetPlugin(ByVal obj As Object)
    Set m_Plugin = obj

    If Not m_Plugin Is Nothing Then
        lblStatus.Caption = "Status: Connected (from host)"
        lblStatus.ForeColor = &H8000&
        RefreshProperties
    Else
        lblStatus.Caption = "Status: Object is Nothing"
        lblStatus.ForeColor = &HFF&
    End If
End Sub

''' <summary>
''' Read all IComPlugin properties via late binding
''' </summary>
Private Sub RefreshProperties()
    If m_Plugin Is Nothing Then Exit Sub

    On Error Resume Next

    ' IComPlugin.Name (DispId 1)
    txtName.Text = m_Plugin.Name

    ' IComPlugin.Version (DispId 2)
    txtVersion.Text = m_Plugin.Version

    ' IComPlugin.Description (DispId 3)
    txtDesc.Text = m_Plugin.Description

    If Err.Number <> 0 Then
        lblStatus.Caption = "Status: Error - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub btnRefresh_Click()
    RefreshProperties
    lblStatus.Caption = "Status: Properties refreshed"
End Sub

''' <summary>
''' Create C# COM object independently via Reg-Free COM
'''
''' Strategy:
'''   1. Try direct CreateObject (works in compiled DLL mode where
'''      parent manifest's <dependency> chain is already active)
'''   2. If that fails (VB6 IDE debug mode), use modActCtx to
'''      manually activate CSharpComPlugin.manifest via
'''      Win32 Activation Context API, then retry CreateObject
''' </summary>
Private Sub btnCreateViaManifest_Click()
    On Error Resume Next

    ' Step 1: Try direct CreateObject
    '   In compiled mode: VB6ComPlugin.manifest -> CSharpComPlugin dependency
    '   SxS activation context is already active from the host process
    Dim obj As Object
    Set obj = CreateObject("ComPluginDemo.Calculator")

    If Err.Number = 0 Then
        ' Success! (compiled DLL mode or COM registered in system)
        GoTo CreateSuccess
    End If

    Err.Clear
    On Error GoTo 0

    ' Step 2: Manual manifest activation (VB6 IDE debug mode)
    '   VB6 IDE process has no SxS activation context,
    '   so we use CreateActCtxW/ActivateActCtx from modActCtx
    Dim sManifest As String
    sManifest = FindCSharpManifest()

    If Len(sManifest) = 0 Then
        lblStatus.Caption = "Status: Cannot find CSharpComPlugin.manifest"
        lblStatus.ForeColor = &HFF&
        Exit Sub
    End If

    If Not ActivateManifestCtx(sManifest) Then
        lblStatus.Caption = "Status: Failed to activate manifest (Err " & CStr(Err.LastDllError) & ")"
        lblStatus.ForeColor = &HFF&
        Exit Sub
    End If

    On Error Resume Next
    Set obj = CreateObject("ComPluginDemo.Calculator")
    If Err.Number <> 0 Then
        lblStatus.Caption = "Status: CreateObject failed - " & Err.Description
        lblStatus.ForeColor = &HFF&
        Err.Clear
        DeactivateManifestCtx
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

CreateSuccess:
    ' Use the newly created COM object
    Set m_Plugin = obj

    On Error Resume Next
    obj.Initialize
    On Error GoTo 0

    lblStatus.Caption = "Status: Created via Manifest (Reg-Free COM)"
    lblStatus.ForeColor = &H8000&
    lblBindInfo.Caption = "Created: CreateObject(""ComPluginDemo.Calculator"") via SxS"
    RefreshProperties
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
        lblStatus.Caption = "Status: Initialize() OK"
    End If
    On Error GoTo 0
End Sub

''' <summary>
''' Call IComPlugin.Execute(input) (DispId 5)
''' </summary>
Private Sub btnExecute_Click()
    If m_Plugin Is Nothing Then
        txtResult.Text = "Error: No plugin. Click 'Create via Manifest' first."
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
        Err.Clear
    Else
        txtResult.Text = sResult
        lblStatus.Caption = "Status: Execute() OK"
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
        lblStatus.Caption = "Status: Shutdown() OK"
    End If
    On Error GoTo 0
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set m_Plugin = Nothing
    ' Clean up activation context if we activated it manually
    If IsManifestActive Then DeactivateManifestCtx
End Sub
