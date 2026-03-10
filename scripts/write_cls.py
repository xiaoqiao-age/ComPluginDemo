"""Write StringProcessor.cls in GBK encoding with CRLF line endings"""
import os

SLN_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLS_PATH = os.path.join(SLN_DIR, "src", "VB6ComPlugin", "StringProcessor.cls")

content = """\
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================
' StringProcessor - VB6 COM Plugin
'
' ProgID: VB6ComPlugin.StringProcessor
' Commands: upper/lower/reverse/len/words/calc/show/explore
'===============================================================================

Option Explicit

Private m_Initialized As Boolean
Private m_Calculator As Object

' Cached Calculator properties (set by host via SetCalculatorObject)
Private m_CalcName As String
Private m_CalcVersion As String
Private m_CalcDescription As String
Private m_CalcObj As Object

'--- IComPlugin Interface ---

Public Property Get Name() As String
    Name = "StringProcessor"
End Property

Public Property Get Version() As String
    Version = "1.0.0"
End Property

Public Property Get Description() As String
    Description = "String Processor (VB6) - Commands: upper/lower/reverse/len/words/calc/show/explore"
End Property

Public Sub Initialize()
    If m_Initialized Then Exit Sub

    ' Try to create C# Calculator COM object
    On Error Resume Next
    Set m_Calculator = CreateObject("ComPluginDemo.Calculator")
    If Not m_Calculator Is Nothing Then
        m_Calculator.Initialize
    End If
    On Error GoTo 0

    m_Initialized = True
End Sub

''' <summary>
''' Called by C# host to pass Calculator COM object reference
''' </summary>
Public Sub SetCalculatorObject(ByVal objCalc As Object)
    On Error Resume Next
    Set m_CalcObj = objCalc
    If Not objCalc Is Nothing Then
        m_CalcName = objCalc.Name
        m_CalcVersion = objCalc.Version
        m_CalcDescription = objCalc.Description
    End If
    On Error GoTo 0
End Sub

''' <summary>
''' Called by C# host to pass the base directory for manifest search.
''' Stored in modActCtx.g_ManifestBaseDir for use by frmComExplorer.
''' </summary>
Public Sub SetManifestBaseDir(ByVal sDir As String)
    g_ManifestBaseDir = sDir
End Sub

Public Function Execute(ByVal sInput As String) As String
    If Not m_Initialized Then Initialize

    Dim cmd As String
    Dim param As String
    Dim spacePos As Long

    sInput = Trim$(sInput)
    If Len(sInput) = 0 Then
        Execute = "Usage: <command> <text>" & vbCrLf & _
                  "Commands: upper, lower, reverse, len, words, calc, show, explore"
        Exit Function
    End If

    ' Split command and parameter
    spacePos = InStr(sInput, " ")
    If spacePos > 0 Then
        cmd = LCase$(Left$(sInput, spacePos - 1))
        param = Mid$(sInput, spacePos + 1)
    Else
        cmd = LCase$(sInput)
        param = ""
    End If

    ' Execute command
    Select Case cmd
        Case "upper"
            If Len(param) = 0 Then
                Execute = "Error: upper requires text"
            Else
                Execute = UCase$(param)
            End If

        Case "lower"
            If Len(param) = 0 Then
                Execute = "Error: lower requires text"
            Else
                Execute = LCase$(param)
            End If

        Case "reverse"
            If Len(param) = 0 Then
                Execute = "Error: reverse requires text"
            Else
                Execute = ReverseString(param)
            End If

        Case "len"
            If Len(param) = 0 Then
                Execute = "Error: len requires text"
            Else
                Execute = "Length: " & CStr(Len(param))
            End If

        Case "words"
            If Len(param) = 0 Then
                Execute = "Error: words requires text"
            Else
                Execute = "Word count: " & CStr(CountWords(param))
            End If

        Case "calc"
            If Len(param) = 0 Then
                Execute = "Error: calc requires expression"
            ElseIf m_Calculator Is Nothing Then
                Execute = "Error: C# Calculator COM not available"
            Else
                On Error Resume Next
                Dim calcResult As String
                calcResult = m_Calculator.Execute(param)
                If Err.Number <> 0 Then
                    Execute = "Calc error: " & Err.Description
                    Err.Clear
                Else
                    Execute = "[via C# Calculator] " & calcResult
                End If
                On Error GoTo 0
            End If

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
            ' Open COM Explorer Form (late binding, all IComPlugin methods)
            If m_CalcObj Is Nothing Then
                Execute = "Error: Calculator object not set. Host must call SetCalculatorObject first."
            Else
                Dim frmEx As frmComExplorer
                Set frmEx = New frmComExplorer
                frmEx.SetPlugin m_CalcObj
                frmEx.Show vbModal
                Set frmEx = Nothing
                Execute = "COM Explorer form closed."
            End If

        Case Else
            Execute = "Unknown command: " & cmd & vbCrLf & _
                      "Commands: upper, lower, reverse, len, words, calc, show, explore"
    End Select
End Function

Public Sub Shutdown()
    If Not m_Calculator Is Nothing Then
        On Error Resume Next
        m_Calculator.Shutdown
        On Error GoTo 0
        Set m_Calculator = Nothing
    End If
    Set m_CalcObj = Nothing
    m_Initialized = False
End Sub

'--- Helper Functions ---

Private Function ReverseString(ByVal s As String) As String
    Dim i As Long
    Dim result As String
    result = ""
    For i = Len(s) To 1 Step -1
        result = result & Mid$(s, i, 1)
    Next i
    ReverseString = result
End Function

Private Function CountWords(ByVal s As String) As Long
    Dim count As Long
    Dim inWord As Boolean
    Dim i As Long
    Dim ch As String

    count = 0
    inWord = False

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            If inWord Then
                count = count + 1
                inWord = False
            End If
        Else
            inWord = True
        End If
    Next i

    If inWord Then count = count + 1

    CountWords = count
End Function
"""

# Write with GBK encoding and CRLF
with open(CLS_PATH, "w", encoding="gbk", newline="\r\n") as f:
    f.write(content)

print(f"[OK] Written: {CLS_PATH}")
print(f"     Size: {os.path.getsize(CLS_PATH)} bytes")
