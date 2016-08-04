# VBA-Dynamic-Array
A sequence container representing an array that can change in size. 

## Getting Started
The following example shows how to declare, populate and access the array.

```vb
Dim v As Vector
Set v = New Vector

v.Push 1
v.Push 2
v.Push 3

Dim i As Long
For i = 0 To v.Size - 1
  Debug.Print v.Element(i) 
Next i
```

Pushing a lot of data can be time consuming. Instead of increasing the array capacity one element at the time it could be a better idea to increase the capacity by larger segments. In the example below it is increased by 10. When you are done populating the array just shrink it to the appropriate size.

```vb
Dim v As Vector
Set v = New Vector

v.AutoAlloc = 10

Dim i As Long
i = 0
Do While some condition
  v.Push i
  i = i + 1
Loop

v.Shrink
```

Converting a Vector to a long array can be done by
```vb
Dim longArr() As Long
longArr  = v.ToLong
```

