Option Explicit

Private data() As Long
Private lastIndex As Long
Private vCapacity As Long
Private alloc As Long

Public Function ToLong() As Long()
    
    Dim longArr() As Long
    
    If lastIndex <> -1 Then
        ReDim longArr(lastIndex)
    
        Dim i As Long
        For i = 0 To lastIndex
            longArr(i) = data(i)
        Next i
    End If
    
    ToLong = longArr

End Function

Public Function Size() As Long
    
    Size = lastIndex + 1
    
End Function

Public Property Get Capacity() As Long

    Capacity = vCapacity
    
End Property

Public Property Let AutoAlloc(value As Long)
    
    alloc = value
    
End Property

Public Property Get AutoAlloc() As Long

    AutoAlloc = alloc
    
End Property

Public Sub Reserve(n As Long)
           
    If n > vCapacity Then
        ReDim Preserve data(n - 1)
        vCapacity = n
    End If

End Sub

Public Sub Realloc(n As Long)
    
    If n = 0 Then
        Delete
    ElseIf n > 0 Then
        If n - 1 < lastIndex Then lastIndex = n - 1
        ReDim Preserve data(n - 1)
        vCapacity = n
    End If
    
End Sub

Public Sub Resize(n As Long)
    
    If Not n < 0 Then
        If n - 1 > lastIndex Then Reserve n
        lastIndex = n - 1
    Else
        Err.Raise 5, "Vector.Resize", "Negative argument is not allowed."
    End If
    
End Sub

Public Sub Shrink()
    
    Realloc lastIndex + 1
    
End Sub

Public Function Element(index As Long) As Long
    
    If index > lastIndex Then Err.Raise 9
    Element = data(index)
    
End Function

Public Sub Push(Element As Long)
     
    If Capacity = 0 Then
        Reserve alloc
    ElseIf lastIndex = Capacity - 1 Then
        Reserve lastIndex + 1 + alloc
    End If
    
    lastIndex = lastIndex + 1
    data(lastIndex) = Element
  
End Sub


Public Function IsEmpty() As Boolean
    
    If Capacity = 0 Then
        IsEmpty = True
    Else
        IsEmpty = False
    End If
        
End Function

Public Function PushArr(arr() As Long)
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        Push arr(i)
    Next i
    
End Function

Public Sub Delete()
    
    Erase data
    lastIndex = -1
    vCapacity = 0
    
End Sub


Private Sub Class_Initialize()
    
    lastIndex = -1
    vCapacity = 0
    alloc = 1
    
End Sub


