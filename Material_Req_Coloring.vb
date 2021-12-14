Option Explicit
Sub Coloring()

Dim lastcolumn As Integer
lastcolumn = Sheet1.UsedRange.Columns.Count
Dim lastrow As Integer



Dim a As Integer
Dim b As Integer
Dim dum As Integer
Dim lum As Integer
Dim dum1 As Integer
Dim bos As Integer
Dim bos1 As Integer
Dim bos2 As Integer
Dim n As Integer
Dim o As Integer
Dim k As Integer
Dim yu As Integer
Dim pp As Integer


Range("E3:E800").UnMerge
lastrow = Sheet1.UsedRange.Rows.Count
n = 5
o = 5

For n = 5 To lastrow

If IsEmpty(Sheet1.Cells(n, o)) Then
Sheet1.Cells(n, o) = Sheet1.Cells(n - 1, o)
End If
Next n



n = 5
o = 5
b = 0
a = 0

yu = 0
Sheet1.Range("A1:AA1000").Font.Size = 12
Sheet1.Range("A1:AA1000").Font.Name = "Calibri"
Sheet1.Range("A1:AA1000").Font.Color = vbBlack


For yu = 1 To lastrow - 4

a = 1
b = 0
Do While Sheet1.Cells(n, o) = Sheet1.Cells(n + a, o) And IsEmpty(Sheet1.Cells(n + a, o)) = False
a = a + 1
b = b + 1

Loop
'MsgBox b


For k = 1 To lastcolumn - 7

pp = 0
If Sheet1.Cells(n, o + 2) = "Eldeki" And Sheet1.Cells(n, o + 2 + k) > 0 Then
bos1 = 0
For bos = 1 To b
If Sheet1.Cells(n + bos, o + 2) = "Teslim almadaki SS" And Sheet1.Cells(n + bos, o + 2 + k) > 0 Then
pp = 1
ElseIf IsEmpty(Sheet1.Cells(n + bos, o + 2 + k)) Then
bos1 = bos1 + 1
End If
Next bos
If bos1 + pp + 1 = b + 1 And pp = 1 Then
For bos = 0 To b
Sheet1.Cells(n + bos, o + 2 + k).Interior.Color = RGB(141, 180, 226)
Next bos
End If
End If



If Sheet1.Cells(n, o + 2) = "Eldeki" And Sheet1.Cells(n, o + 2 + k) > 0 Then

dum1 = 0
For dum = 1 To b Step 1
If IsEmpty(Sheet1.Cells(n + dum, o + 2 + k)) Then
dum1 = dum1 + 1
ElseIf Sheet1.Cells(n + dum, o + 2) = "Eldeki" And Sheet1.Cells(n + dum, o + 2 + k) > 0 Then
dum1 = dum1 + 1
End If
Next dum
If dum1 = b Then
For dum = 0 To b
Sheet1.Cells(n + dum, o + 2 + k).Interior.Color = RGB(192, 239, 206)
Next dum
End If
End If

If Sheet1.Cells(n, o + 2) = "Teslim almadaki SS" And Sheet1.Cells(n, o + 2 + k) > 0 Then

dum1 = 0
For dum = 1 To b Step 1

If Sheet1.Cells(n + dum, o + 2) = "Teslim almadaki SS" And Sheet1.Cells(n + dum, o + 2 + k) > 0 Then
dum1 = dum1 + 1
End If
Next dum
If dum1 = b Then
For dum = 0 To b
Sheet1.Cells(n + dum, o + 2 + k).Interior.Color = RGB(141, 180, 226)
Next dum
End If
End If

bos2 = 0
dum1 = 0
If IsEmpty(Sheet1.Cells(n, o + 2 + k)) Then
For dum = 1 To b
If IsEmpty(Sheet1.Cells(n + dum, o + 2 + k)) Then
dum1 = dum1 + 1
bos2 = bos2 + 1
ElseIf Sheet1.Cells(n + dum, o + 2) = "Eldeki" And Sheet1.Cells(n + dum, o + 2 + k) > 0 Then
dum1 = dum1 + 1
ElseIf Sheet1.Cells(n + dum, o + 2) = "Teslim almadaki SS" And Sheet1.Cells(n + dum, o + 2 + k) > 0 Then
bos2 = bos2 + 1
End If

Next dum
If dum1 = b Then
For dum = 0 To b
Sheet1.Cells(n + dum, o + 2 + k).Interior.Color = RGB(192, 239, 206)
Next dum
ElseIf bos2 = b Then
For dum = 0 To b
Sheet1.Cells(n + dum, o + 2 + k).Interior.Color = RGB(141, 180, 226)
Next dum
End If
End If




Next k
'MsgBox dum1

n = n + b + 1

Next yu



End Sub