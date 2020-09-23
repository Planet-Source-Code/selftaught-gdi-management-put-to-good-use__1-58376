VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcGDIObjectStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcGDIObjectStore.cls                   1/17/04
'
'           GENERAL PURPOSE:
'               Public interface consists of only three functions.  Init, which accepts only
'               gdiFont, gdiPen or gdiBrush. AddRef, which takes a pointer to a LOGFONT,
'               LOGPEN or LOGBRUSH structure and returns a hObject.  The structure passed
'               to this function must agree with the constant passed to Init.  Finally, the
'               Release function takes a hObject and decrements the reference count or
'               destroys the gdi object.
'
'               Store each unique GDI object using a hash table for fast lookup and
'               a linked list for quick allocation/deallocation.
'
'               Two linked lists are used, but the nodes contained are the same.  One
'               is used for lookup by data (LOGFONT, LOGBRUSH, LOGPEN), and the other
'               for lookup by handle.
'
'               Each node is a set of 32-bit values followed by the LOG* Structure from
'               above.  For the meaning of the value located at each index see
'               the NODE_* constants near the start of the first function.
'
'           COMPILER SWITCHES:
'               bDebugAddRefReleases - Display debug messages showing incrementing/decrementing reference count.
'
'               bDebugStatistics     -   expose the Statistics function
'
'                                    Display Debug messages if any of the following are leaked:
'               bDebugMemory         -   memory allocated for the linked list storage
'               bDebugBrushes        -   brush handles
'               bDebugPens           -   pen handles
'               bDebugFonts          -   font handles
'
'==================================================================================================

Option Explicit

#Const bDebug = False
#Const bDebugStatistics = True

#Const bDebugAddRefReleases = bDebug

#Const bDebugMemory = bDebug
#Const bDebugBrushes = bDebug
#Const bDebugPens = bDebug
#Const bDebugFonts = bDebug

#If bDebugMemory Then
    Private moDebugMemory As pcDebug
#End If
#If bDebugBrushes Then
    Private moDebugBrushes As pcDebug
#End If
#If bDebugPens Then
    Private moDebugPens As pcDebug
#End If
#If bDebugFonts Then
    Private moDebugFonts As pcDebug
#End If

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

Private Declare Function CreateBrushLong Lib "gdi32.dll" Alias "CreateBrushIndirect" (ByVal lpLogBrush As Long) As Long
Private Declare Function CreateFontLong Lib "gdi32.dll" Alias "CreateFontIndirectA" (ByVal lpLogFont As Long) As Long
Private Declare Function CreatePenLong Lib "gdi32.dll" Alias "CreatePenIndirect" (ByVal lpLogPen As Long) As Long
Private Declare Function DeleteObjectApi Lib "gdi32.dll" Alias "DeleteObject" (ByVal hObject As Long) As Long

Private miDataLen               As Long         'Length of the LOG* structure being used by this instance of the object.

Private miDataTable(0 To 255)   As Long         'Hash table that is indexed by hashing the LOG* structures.
Private miHandleTable(0 To 255) As Long         'Hash table that is indexed by hashing the hObjects.

Private Const NODE_hObj As Long = ZeroL
Private Const NODE_RefCount As Long = 4&
Private Const NODE_NextPtrData As Long = 8&
Private Const NODE_DataTableIndex As Long = 12&
Private Const NODE_NextPtrHandle As Long = 16&
Private Const NODE_Data As Long = 20&           'The total node length is this value + miDataLen bytes.

Private miObjectType            As eGdiObjectType 'only gdiBrush, gdiFont, gdiPen

Private Function pFindNode(ByVal iTableIndex As Long, ByVal iPtr As Long) As Long
    'Search for a node by data (LOGFONT, LOGPEN, LOGBRUSH).  The hash value has
    'already been calculated and is passed in through iTableIndex
    'returns a pointer to the node that was found, or 0 otherwise
    
    pFindNode = miDataTable(iTableIndex)                        'Start with the table value
    
    Do Until pFindNode = ZeroL                                  'if no node, we're done
                                                        
        If MemCmp(UnsignedAdd(pFindNode, NODE_Data), iPtr, miDataLen) = ZeroL _
            Then Exit Function                                  'if the LOG* structures are identical, we've found it.
        pFindNode = MemOffset32(pFindNode, NODE_NextPtrData)    'if that wasn't it, then we'll try the next one.
        
    Loop
    
End Function

Private Function pAddNode(ByVal iDataTableIndex As Long, ByVal iPtr As Long) As Long
    'Add a node.  iDataTableIndex is the hash value of the LOG* structure.  iPtr is a pointer to same.
    'This function is only called when it is determined that an identical gdi object is not already
    'created.  Returns a hObject.
    
    Dim liHandleTableIndex As Long
    Dim liNodePtr As Long
    liNodePtr = GlobalAlloc(ZeroL, NODE_Data + miDataLen)           'allocate the node
    
    If liNodePtr Then                                               'proceed if that succeeded
    
        #If bDebugMemory Then
            If moDebugMemory Is Nothing Then                        'If we're debugging
                Set moDebugMemory = New pcDebug                         'make sure we're init.
                moDebugMemory.Name = "Mem Handle"
                moDebugMemory.Module = "pcGDIObjectStore"
            End If
            moDebugMemory.Add liNodePtr                                 'track this handle
        #End If
        
        If miObjectType = gdiBrush Then
            pAddNode = CreateBrushLong(iPtr)                        'Create a Brush
            #If bDebugBrushes = True Then
                If moDebugBrushes Is Nothing Then                   'If we're debugging
                    Set moDebugBrushes = New pcDebug                    'make sure we're init.
                    moDebugBrushes.Module = "pcGDIObjectStore"
                    moDebugBrushes.Name = "Brush"
                End If                                                  'if the brush was created then track it
                If pAddNode Then moDebugBrushes.Add pAddNode
            #End If
            
        ElseIf miObjectType = gdiFont Then
            pAddNode = CreateFontLong(iPtr)                         'Create a Font
            #If bDebugFonts = True Then
                If moDebugFonts Is Nothing Then                     'If we're debugging
                    Set moDebugFonts = New pcDebug                      'make sure we're init.
                    moDebugFonts.Module = "pcGDIObjectStore"
                    moDebugFonts.Name = "Font"
                End If
                If pAddNode Then moDebugFonts.Add pAddNode              'if the font was created then track it
            #End If
        ElseIf miObjectType = gdiPen Then
            pAddNode = CreatePenLong(iPtr)                          'Create a Pen
            #If bDebugPens = True Then
                If moDebugPens Is Nothing Then                      'If we're debugging
                    Set moDebugPens = New pcDebug
                    moDebugPens.Module = "pcGDIObjectStore"             'make sure we're init.
                    moDebugPens.Name = "Pen"
                End If
                If pAddNode Then moDebugPens.Add pAddNode               'if the pen was created then track it
            #End If
        End If
        
        Debug.Assert pAddNode
        
        If pAddNode Then                                                        'if the object was created
            
            liHandleTableIndex = HashLong(pAddNode)                                 'hash the hObject
            
            MemOffset32(liNodePtr, NODE_RefCount) = OneL                            'store the reference count
            MemOffset32(liNodePtr, NODE_hObj) = pAddNode                            'store the hObject
            MemOffset32(liNodePtr, NODE_DataTableIndex) = iDataTableIndex           'store the Hash of the data to save calling it again.
            MemOffset32(liNodePtr, NODE_NextPtrData) = miDataTable(iDataTableIndex) 'store the pointer to the next node on this index in the data table.
            MemOffset32(liNodePtr, NODE_NextPtrHandle) = miHandleTable(liHandleTableIndex) 'store the pointer to the next node on this index in the handle table.
            CopyMemory ByVal UnsignedAdd(liNodePtr, NODE_Data), ByVal iPtr, miDataLen 'store the LOG* structure
            
            miDataTable(iDataTableIndex) = liNodePtr                                'store this node at the first index in both
            miHandleTable(liHandleTableIndex) = liNodePtr                           'the data table and the handle table.
        Else
            GlobalFree liNodePtr                                                'if the object was not created, free the node
            
        End If
    End If
End Function

Private Function pRelease(ByVal hObject As Long) As Boolean
    'returns a success code
    'decrements reference count or destroy the object and release the node.
    
    Dim liLastNodePtr As Long
    Dim liNodePtr As Long
    Dim liHandleTableIndex As Long
    
    liHandleTableIndex = HashLong(hObject)                                  'hash the hObject
    
    liNodePtr = miHandleTable(liHandleTableIndex)                           'get the first node at this index
    
    Do Until liNodePtr = ZeroL                                              'loop until no more nodes
        
        If MemOffset32(liNodePtr, NODE_hObj) = hObject Then                 'if the hObject matches then
            pRelease = pReleaseNode(hObject, liNodePtr, liHandleTableIndex, liLastNodePtr) 'release the node
            Exit Do                                                                        'and get out
            
        End If
        
        liLastNodePtr = liNodePtr                                           'remember the last node so we can pass it to pReleaseNode
        liNodePtr = MemOffset32(liNodePtr, NODE_NextPtrHandle)              'get a handle to the next node
        
    Loop
    
    Debug.Assert liNodePtr <> ZeroL                                         'if no node, why am i calling?
    
End Function

Private Function pReleaseNode(ByVal hObject As Long, ByVal hNode As Long, ByVal iHandleTableIndex As Long, ByVal hNodePrevious As Long) As Boolean
    'decrement the reference count, delete the hObject and remove the node from the linked list as necessary.
    
    Dim liRefCount As Long
    
    liRefCount = MemOffset32(hNode, NODE_RefCount) - OneL                   'store the new reference count locally
    If liRefCount < ZeroL Then liRefCount = ZeroL: Debug.Assert False       'should never happen!
    MemOffset32(hNode, NODE_RefCount) = liRefCount                          'store the new reference count in the node
            
    If liRefCount = ZeroL Then                                              'if we're all done with this object
        
        pReleaseNode = CBool(DeleteObjectApi(hObject))                      'delete it
        
        Debug.Assert pReleaseNode                                           'if you get this assertion, the most likely scenario is that the object is selected into a DC.
        
        If pReleaseNode Then
            
            #If bDebugBrushes Or bDebugFonts Or bDebugPens Then             'if we're debugging
                
                If miObjectType = gdiBrush Then                             'stop tracking the object
                    #If bDebugBrushes = True Then
                        moDebugBrushes.Remove hObject
                    #End If
                ElseIf miObjectType = gdiFont Then
                    #If bDebugPens = True Then
                        moDebugFonts.Remove hObject
                    #End If
                ElseIf miObjectType = gdiPen Then
                    #If bDebugFonts = True Then
                        moDebugPens.Remove hObject
                    #End If
                End If
                
            #End If
            
            If hNodePrevious = ZeroL Then                                   'if there is no previous node
                miHandleTable(iHandleTableIndex) = MemOffset32(hNode, NODE_NextPtrHandle)   'the first node in this index is now the node after the node we're deleting.
                
            Else                                                            'if there is a previous node
                MemOffset32(hNodePrevious, NODE_NextPtrHandle) _
                  = MemOffset32(hNode, NODE_NextPtrHandle)                  'the previous node's next handle index will now point to the node after the node we're deleting.
                
            End If
            
                                                                            'OK, Halfway done!  the miHandleTable array maintenance is complete
                                                                            'now we begin the miDataTable array maintenance
            
            Dim liDataTableIndex As Long
            Dim liDataNodePtr As Long
            Dim liLastNodePtr As Long
            
            liDataTableIndex = MemOffset32(hNode, NODE_DataTableIndex)      'get the hash value of the LOG* structure.  This could be calculated again here, but we're storing it in the node instead.
            liDataNodePtr = miDataTable(liDataTableIndex)                   'get the first node at this index
            
            Do Until liDataNodePtr = ZeroL                                  'loop while there are nodes.
                
                If liDataNodePtr = hNode Then                               'if we've found the node that we're deleting
                    If liLastNodePtr = ZeroL Then                           'if there is no previous pointer
                        miDataTable(liDataTableIndex) = MemOffset32(liDataNodePtr, NODE_NextPtrData) 'the first node at this index is now the node after the node we just deleted.
                        
                    Else                                                    'if there is a previous pointer
                        MemOffset32(liLastNodePtr, NODE_NextPtrData) _
                          = MemOffset32(liDataNodePtr, NODE_NextPtrData)    'the nextptr member of the previous pointer must now point to the nextptr member of the node we're deleting.
                  
                    End If
                    
                    Exit Do                                                 'found it, we're gone.
                End If
                
                liLastNodePtr = liDataNodePtr                               'remember the last pointer
                liDataNodePtr = MemOffset32(liDataNodePtr, NODE_NextPtrData) 'get the next pointer
                
            Loop
            
            Debug.Assert liDataNodePtr <> ZeroL                             'if there's no node then we're in trouble, the arrays are inconsistent with each other!
            
            GlobalFree hNode                                                'Finally ... free the node
            
            #If bDebugMemory Then
                moDebugMemory.Remove hNode                                  'if we're debugging memory, stop tracking this handle.
            #End If
        End If
    Else
        pReleaseNode = True                                                 'If somebody still wants this object, return success
        
        #If bDebugAddRefReleases Then
            Debug.Print Switch(miObjectType = gdiFont, "Font", _
                               miObjectType = gdiPen, "Pen", _
                               miObjectType = gdiBrush, "Brush") & _
                               " Release Handle: " & MemOffset32(hNode, NODE_hObj) & _
                               " RefCount: " & MemOffset32(hNode, NODE_RefCount)
        #End If
        
        
    End If

End Function


Public Sub Init(ByVal iObjectType As eGdiObjectType)
    'This procedure must be called before any others, can only be called once,
    'and must be passed one of these values: gdiBrush, gdiFont, gdiPen
    
    Static bInit As Boolean
    
    If bInit Then
        Debug.Assert False 'Can only init once!
        Exit Sub
    End If
    
    If iObjectType = gdiFont Then       'store the length of the data we're working with
        Dim ltFont As LOGFONT
        miDataLen = LenB(ltFont)
    ElseIf iObjectType = gdiPen Then
        Dim ltPen As LOGPEN
        miDataLen = LenB(ltPen)
    ElseIf iObjectType = gdiBrush Then
        Dim ltBrush As LOGBRUSH
        miDataLen = LenB(ltBrush)
    Else
        Debug.Assert False
        Exit Sub
    End If
    miObjectType = iObjectType          'store the type of data
    
    bInit = True
    
End Sub

Public Function AddRef(ByVal iPtr As Long) As Long
    'accepts a LOG* structure, depending on what was passed to Init
    'Returns the GDI Object handle
    
    Dim liIndex As Long
    Dim lpNode As Long
    
    If (miDataLen > ZeroL) Then    'make sure we are initialized
        liIndex = Hash(iPtr, miDataLen) 'hash the LOG* structure
        lpNode = pFindNode(liIndex, iPtr) 'find the object if it's already been created.
        If lpNode Then                    'if the object has already been created
            AddRef = MemOffset32(lpNode, NODE_hObj) 'get the hObject
            MemOffset32(lpNode, NODE_RefCount) = MemOffset32(lpNode, NODE_RefCount) + OneL 'inc the ref count
            
            #If bDebugAddRefReleases Then
                Debug.Print Switch(miObjectType = gdiFont, "Font", _
                                   miObjectType = gdiPen, "Pen", _
                                   miObjectType = gdiBrush, "Brush") & _
                                   " AddRef Handle:  " & MemOffset32(lpNode, NODE_hObj) & _
                                   " RefCount: " & MemOffset32(lpNode, NODE_RefCount)
            #End If
            
        Else    'the object has not already been created
            AddRef = pAddNode(liIndex, iPtr)
        End If
    End If
End Function

Public Function Release(ByVal hObject As Long) As Long
    If miDataLen > ZeroL Then
        'return 1 for success, 0 for failure.  This emulates the api call DeleteObject.
        Release = -pRelease(hObject)
    End If
End Function

#If bDebugStatistics Then
    Public Sub Statistics(ByRef iRequested As Long, ByRef iCreated As Long)
        Dim i As Long
        Dim hNode As Long
        
        iRequested = ZeroL
        iCreated = ZeroL
        
        For i = LBound(miHandleTable) To UBound(miHandleTable)
            hNode = miHandleTable(i)
            Do Until hNode = ZeroL
                iCreated = iCreated + OneL
                iRequested = iRequested + MemOffset32(hNode, NODE_RefCount)
                hNode = MemOffset32(hNode, NODE_NextPtrHandle)
            Loop
        Next
    End Sub
#End If
