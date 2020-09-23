VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcArrayPtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcArrayPtr.bas                      1/19/04
'
'           GENERAL PURPOSE:
'               Replace the SAFEARRAY Pointer of a VB array with our own,
'               forcing it to access the memory we point it to.
'
'           LINEAGE:
'               Numerous submission from many authors on www.pscode.com
'
'           COMPILER SWITCHES:
'               bVBVMTypeLib - Toggles dependence on VBVM6Lib.tlb
'
'==================================================================================================

Option Explicit

#Const bVBVMTypeLib = False

#If bVBVMTypeLib = False Then
    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
#End If

Private Type tSafeArray1D
    Dimensions As Integer
    Attributes As Integer
    BytesPerElement As Long
    Locks As Long
    DataPointer As Long
    Elements As Long
    LBound As Long
End Type

#If Not bVBVMTypeLib Then
    Private Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
#End If

Private mtHeader As tSafeArray1D
Private miOldHeader As Long
Private miArrPtr As Long

Public Sub SetArrayByte(ByRef yArray() As Byte)
    Destroy
    
    Const SAFEARRAY_AUTO = &H1
    Const SAFEARRAY_FIXEDSIZE = &H10
    
    With mtHeader
        .Dimensions = OneL                                          'initialize the structure
        .Attributes = SAFEARRAY_AUTO Or SAFEARRAY_FIXEDSIZE
        .BytesPerElement = OneL
        .Locks = ZeroL
        '.DataPointer = ZeroL
        '.Elements = ZeroL
        .LBound = ZeroL
    End With
    
    miArrPtr = ArrPtr(yArray)
    #If bVBVMTypeLib Then
        miOldHeader = MemLong(ByVal miArrPtr)
        Debug.Assert miOldHeader = ZeroL
        MemLong(ByVal miArrPtr) = VarPtr(mtHeader)
    #Else
        CopyMemory miOldHeader, ByVal miArrPtr, 4&                      'store the current array pointer
        Debug.Assert miOldHeader = ZeroL                                'should be undefined
        CopyMemory ByVal miArrPtr, VarPtr(mtHeader), 4&                 'replace the current pointer
    #End If

    
End Sub


Public Sub SetArrayLong(ByRef iArray() As Long)
    Destroy
    
    Const SAFEARRAY_AUTO = &H1
    Const SAFEARRAY_FIXEDSIZE = &H10
    
    With mtHeader
        .Dimensions = OneL                                          'initialize the structure
        .Attributes = SAFEARRAY_AUTO Or SAFEARRAY_FIXEDSIZE
        .BytesPerElement = 4&
        .Locks = ZeroL
        '.DataPointer = ZeroL
        '.Elements = ZeroL
        .LBound = ZeroL
    End With
    
    miArrPtr = ArrPtr(iArray)
    #If bVBVMTypeLib Then
        miOldHeader = MemLong(ByVal miArrPtr)
        Debug.Assert miOldHeader = ZeroL
        MemLong(ByVal miArrPtr) = VarPtr(mtHeader)
    #Else
        CopyMemory miOldHeader, ByVal miArrPtr, 4&                      'store the current array pointer
        Debug.Assert miOldHeader = ZeroL                                'should be undefined
        CopyMemory ByVal miArrPtr, VarPtr(mtHeader), 4&                 'replace the current pointer
    #End If
    
End Sub

Public Sub Destroy()
    If CBool(miArrPtr) Then
        #If bVBVMTypeLib Then
            MemLong(ByVal miArrPtr) = miOldHeader
        #Else
            CopyMemory ByVal miArrPtr, miOldHeader, 4&
        #End If
        
        miArrPtr = ZeroL
        miOldHeader = ZeroL
    End If
End Sub

Public Function PointToString(ByRef s As String) As Long
    mtHeader.DataPointer = StrPtr(s)
    PointToString = LenB(s)
    mtHeader.Elements = PointToString
End Function

Public Sub PointToLong(ByRef i As Long)
    mtHeader.DataPointer = VarPtr(i)
    mtHeader.Elements = 4&
End Sub

Public Sub Point(ByVal iPtr As Long, ByVal iLen As Long)
    mtHeader.DataPointer = iPtr
    mtHeader.Elements = iLen
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub


