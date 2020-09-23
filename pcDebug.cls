VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'cDebugHandles.bas                      12/15/04
'
'           GENERAL PURPOSE:
'               Track menu, string, memory, gdi and other handles for debugging purposes.
'
'           LINEAGE:
'               N/A
'
'==================================================================================================

Option Explicit

Public Name As String
Public Module As String
Public PrintAddRemove As Boolean

Private mColl As New Collection

Public Sub Add(ByVal h As Long)
Debug.Assert LenB(Name)
    On Error GoTo catch
    mColl.Add h, CStr(h)
    If PrintAddRemove Then
        Debug.Print "Add " & Name & " " & h
    End If
    Exit Sub
catch:
    pErr
End Sub

Public Sub Remove(ByVal h As Long)
Debug.Assert LenB(Name)
    On Error GoTo catch
    mColl.Remove CStr(h)
    If PrintAddRemove Then
        Debug.Print "Del " & Name & " " & h
    End If
    Exit Sub
catch:
    pErr
End Sub

Private Sub pErr()
    Debug.Print "Debug Err: " & Err.Number, Err.Description
    Debug.Assert False
End Sub

Private Sub Class_Terminate()
    If Not mColl Is Nothing Then
        If mColl.Count Then
'            Dim ls As String
'            Dim v As Variant
'            For Each v In mColl
'                ls = ls & v & vbCrLf
'            Next
            Debug.Print "Leaked " & mColl.Count & " " & Name & "(s) from " & Module
        End If
    End If
    Debug.Print "Term " & mColl.Count & " " & Name, Module
End Sub
