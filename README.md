<div align="center">

## A Very Simple Sort


</div>

### Description

A very simpile example of sorting. Not the most efficeint, but easy for beginners to see whats happening.
 
### More Info
 
Unsorted Array, Sort order (True = Sort Ascending)

Sorted Array


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Paul Crowdy](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-crowdy.md)
**Level**          |Beginner
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/paul-crowdy-a-very-simple-sort__1-38375/archive/master.zip)





### Source Code

```
Private Sub Form_Load()
  Dim a(10)
  a(0) = 2
  a(1) = 5
  a(2) = 7
  a(3) = 6
  a(4) = 13
  a(5) = "b"
  a(6) = 65
  a(7) = 0
  a(8) = 4
  a(9) = "a"
  a(10) = 1000
  Sort a, False
  For i = 0 To 10
    Debug.Print a(i)
  Next i
End Sub
Sub Sort(ByRef Arr() As Variant, Optional ByVal bAsc As Boolean = True)
  Dim Done As Boolean
  Done = False
  Do While Done = False
    Done = True
    For i = 0 To UBound(Arr) - 1
      If (Arr(i) > Arr(i + 1) And bAsc) Or (Arr(i) < Arr(i + 1) And Not bAsc) Then Swap Arr(i), Arr(i + 1): Done = False
    Next i
  Loop
End Sub
Sub Swap(ByRef a As Variant, ByRef b As Variant)
  Dim tmp As Variant
  tmp = a
  a = b
  b = tmp
End Sub
```

