Attribute VB_Name = "modActCtx"
Option Explicit

' ============================================================
' Win32 Activation Context API for Reg-Free COM
'
' VB6 IDE debug mode: VB6 process has no SxS activation context,
' so CreateObject("ComPluginDemo.Calculator") would fail.
' This module manually activates CSharpComPlugin.manifest
' to enable Reg-Free COM object creation.
'
' Compiled DLL mode: Not needed (parent manifest dependency chain
' handles it), but harmless to call.
' ============================================================

Private Type ACTCTX
    cbSize As Long
    dwFlags As Long
    lpSource As Long
    wProcessorArchitecture As Integer
    wLangId As Integer
    lpAssemblyDirectory As Long
    lpResourceName As Long
    lpApplicationName As Long
    hModule As Long
End Type

Private Const ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID As Long = &H4
Private Const INVALID_HANDLE_VALUE As Long = -1

Private Declare Function CreateActCtxW Lib "kernel32" (pActCtx As ACTCTX) As Long
Private Declare Function ActivateActCtx Lib "kernel32" ( _
    ByVal hActCtx As Long, lpCookie As Long) As Long
Private Declare Function DeactivateActCtx Lib "kernel32" ( _
    ByVal dwFlags As Long, ByVal ulCookie As Long) As Long
Private Declare Sub ReleaseActCtx Lib "kernel32" (ByVal hActCtx As Long)

' Module-level activation context state
Private m_hActCtx As Long
Private m_Cookie As Long
Private m_Active As Boolean

' Host-provided base directory for finding manifests
Public g_ManifestBaseDir As String

''' <summary>
''' Find CSharpComPlugin.manifest by searching known locations
''' </summary>
Public Function FindCSharpManifest() As String
    Dim sPaths(3) As String
    Dim i As Long

    ' Path 0: Host-provided base dir (C# host passes its BaseDir)
    If Len(g_ManifestBaseDir) > 0 Then
        sPaths(0) = g_ManifestBaseDir & "CSharpComPlugin\CSharpComPlugin.manifest"
    End If

    ' Path 1: App.Path\CSharpComPlugin\ (compiled DLL loaded by host)
    sPaths(1) = App.Path & "\CSharpComPlugin\CSharpComPlugin.manifest"

    ' Path 2: App.Path\VB6ComPlugin\CSharpComPlugin\ (host with VB6 subdir)
    sPaths(2) = App.Path & "\VB6ComPlugin\CSharpComPlugin\CSharpComPlugin.manifest"

    ' Path 3: CurDir (VB6 IDE may set CWD to project dir)
    sPaths(3) = CurDir$ & "\CSharpComPlugin\CSharpComPlugin.manifest"

    For i = 0 To 3
        If Len(sPaths(i)) > 0 Then
            If Len(Dir$(sPaths(i))) > 0 Then
                FindCSharpManifest = sPaths(i)
                Exit Function
            End If
        End If
    Next i

    FindCSharpManifest = ""
End Function

''' <summary>
''' Activate a manifest using Win32 Activation Context API.
''' After activation, CreateObject can resolve ProgIDs declared in the manifest.
''' </summary>
Public Function ActivateManifestCtx(ByVal sManifestPath As String) As Boolean
    If m_Active Then
        ActivateManifestCtx = True
        Exit Function
    End If

    Dim ctx As ACTCTX
    ctx.cbSize = 32  ' sizeof(ACTCTX) on 32-bit
    ctx.lpSource = StrPtr(sManifestPath)

    ' Set assembly directory so SxS can find dependent DLLs
    Dim sDir As String
    Dim pos As Long
    pos = InStrRev(sManifestPath, "\")
    If pos > 0 Then
        sDir = Left$(sManifestPath, pos - 1)
        ctx.dwFlags = ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID
        ctx.lpAssemblyDirectory = StrPtr(sDir)
    End If

    m_hActCtx = CreateActCtxW(ctx)
    If m_hActCtx = INVALID_HANDLE_VALUE Then
        ActivateManifestCtx = False
        Exit Function
    End If

    If ActivateActCtx(m_hActCtx, m_Cookie) = 0 Then
        ReleaseActCtx m_hActCtx
        m_hActCtx = 0
        ActivateManifestCtx = False
        Exit Function
    End If

    m_Active = True
    ActivateManifestCtx = True
End Function

''' <summary>
''' Deactivate and release the activation context
''' </summary>
Public Sub DeactivateManifestCtx()
    If Not m_Active Then Exit Sub
    DeactivateActCtx 0, m_Cookie
    ReleaseActCtx m_hActCtx
    m_hActCtx = 0
    m_Cookie = 0
    m_Active = False
End Sub

''' <summary>
''' Check if an activation context is currently active
''' </summary>
Public Property Get IsManifestActive() As Boolean
    IsManifestActive = m_Active
End Property
