Option Explicit

Private data() As Long
Private lastIndex As Long
Private vCapacity As Long
Private alloc As Long

'''
''' Converts the vector to a long array
'''
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

'''
''' Returns the number of elements in the vector
'''
Public Function Size() As Long
    
    Size = lastIndex + 1
    
End Function

'''
''' Returns the size of the storage space currently allocated for the vector
'''
Public Function Capacity() As Long

    Capacity = vCapacity
    
End Function

Public Property Let AutoAlloc(value As Long)
    
    alloc = value
    
End Property

Public Property Get AutoAlloc() As Long

    AutoAlloc = alloc
    
End Property

''' Explicitly allocate storage space for n elements
''' This Sub does not affect the size of the vector nor alter its content
''' Nothing happens if n is smaller than the size of the vector
Public Sub Reserve(n As Long)
           
    If n > vCapacity Then
        ReDim Preserve data(n - 1)
        vCapacity = n
    End If

End Sub

''' Explicitly reallocate the storage space of the vector
''' If n is smaller than the size of the vector all elements beyond subscript n - 1 will be removed
''' In all other cases the size of the vector is not affected
''' Calling Realloc with n = 0 is equivalent to Delete
Public Sub Realloc(n As Long)
    
    If n = 0 Then
        Delete
    ElseIf n > 0 Then
        If n - 1 < lastIndex Then lastIndex = n - 1
        ReDim Preserve data(n - 1)
        vCapacity = n
    Else
        Err.Raise 5, "Vector.Resize", "Negative argument is not allowed."
    End If
    
End Sub

''' Changes the size of the vector
''' If n is greater than the current size, all new elements will be set to zero
''' The storage space is reallocated automatically if needed
''' If n is less than the current size, all elements beyond subscript n - 1 will be removed
Public Sub Resize(n As Long)
    
    If Not n < 0 Then
        If n - 1 > lastIndex Then
            Reserve n
        ElseIf n - 1 < lastIndex Then
            Dim i As Long
            For i = n To lastIndex
                data(i) = 0
            Next i
        End If
        lastIndex = n - 1
    Else
        Err.Raise 5, "Vector.Resize", "Negative argument is not allowed."
    End If
    
End Sub

'''
''' Reduces the capacity to fit the size of the vector
'''
Public Sub Shrink()
    
    Realloc lastIndex + 1
    
End Sub

'''
''' Access an element in the vector
'''
Public Function Element(index As Long) As Long
    
    If index > lastIndex Then Err.Raise 9
    Element = data(index)
    
End Function

'''
''' Adds an element to the back of the vector
''' Automatically increases the capacity according to AutoAlloc if needed
'''
Public Sub Push(Element As Long)
     
    If Capacity = 0 Then
        Reserve alloc
    ElseIf lastIndex = Capacity - 1 Then
        Reserve lastIndex + 1 + alloc
    End If
    
    lastIndex = lastIndex + 1
    data(lastIndex) = Element
  
End Sub

'''
''' Returns true if the vector is empty
'''
Public Function IsEmpty() As Boolean
    
    If Capacity = 0 Then
        IsEmpty = True
    Else
        IsEmpty = False
    End If
        
End Function


'''
''' Adds an array to the back of the vector
''' Automatically increases the capacity according to AutoAlloc if needed
'''
Public Function PushArr(arr() As Long)
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        Push arr(i)
    Next i
    
End Function

'''
''' Deletes all elements in the vector and deallocates the storage space
'''
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


