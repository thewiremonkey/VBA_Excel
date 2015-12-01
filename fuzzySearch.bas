Attribute VB_Name = "FuzzySearch"

Sub CloseMatchWhole()
'Application.ScreenUpdating = False
Dim Criteria() As Variant
Dim Exceptions() As Variant
Dim ws As Worksheet
Dim counter As Long 'counts cells
Dim i, j, k, n, m, p As Long 'i counts whole search criteria, k counts partial search criteria_
'n counts misses after a hit
Dim st, stF, stO As String
Dim s As Variant
Dim x, y, z As Long 'x counts hits, y is the length of the search criteria
Dim r As Range
Dim rCrit As Range
Dim rE As Range

Dim start, finish
start = Timer


Set r = Application.InputBox(prompt:="Please Select Range", Title:="Range Select", Type:=8)
Set rCrit = Application.InputBox(prompt:="Please Select Criteria Range", Title:="Range Select", Type:=8)

Debug.Print r.Address
Debug.Print rCrit.Address


ReDim Criteria(0 To 1)
        For i = 0 To rCrit.Cells.Count - 1
        ReDim Preserve Criteria(0 To i)
        Criteria(i) = rCrit.Cells(i + 1).Value
        Next i



z = 0
p = 0
m = 0
n = 0
x = 0


For counter = 1 To r.Cells.Count
st = r.Cells(counter).Value
For i = 0 To UBound(Criteria)
        
    If Not Criteria(i) = "" Then
    
    
    
    On Error Resume Next
        y = Len(Criteria(i))
        
        For k = 1 To y - 2
        
            stF = Mid(Criteria(i), k, 3)
                Select Case p
                Case 0
                m = InStr(1, st, stF, vbBinaryCompare)
                p = m
                
                Case 1
                m = InStr(p, st, stF, vbBinaryCompare)
                p = p
                
                Case Is >= 2
                m = InStr(p - 1, st, stF, vbBinaryCompare)
                p = p
                End Select
                
                If m > 0 Then
                p = p
                x = x + 1
                
                    Else:
                    
                    n = n + 1
                        If n < 5 And (m - p) > k + 4 Then
                        p = 0
                        Else:
                        p = p
                        
                    End If
                
                End If

                
                
                    If x > y * 0.45 And x < y * 0.6 And p <> 0 Then
                        z = z + 1
                        If (Mid(r.Cells(counter), p, 1)) = UCase(Mid(r.Cells(counter), p, 1)) Then
                            With r.Cells(counter).Characters(p, Len(Criteria(i))).Font
                            .color = RGB(0, 102, 0)
                            .Bold = True
                            End With
                        End If
                        
                    End If
                        
                    If x >= y * 0.7 And Not p = 0 Then
                        z = z + 1
                        If (Mid(r.Cells(counter), p, 1)) = UCase(Mid(r.Cells(counter), p, 1)) Then
                        r.Cells(counter).Characters(p, Len(Criteria(i))).Font.color = RGB(0, 37, 147)
'                        stO = Mid(r.Cells(counter), p, Len(Criteria(i)))
'                            If Not stO = "" Then
'                                ActiveWorkbook.Sheets(1).Range("a1").Activate
'                                ActiveCell.Offset(counter, 0).Value = stO
'                                ActiveCell.Offset(counter, 1).Value = Criteria(i)
'                                ActiveCell.Offset(counter, 2).Value = r.Cells(counter).Address
'                            End If
                        
                        
                        End If

                    End If
        Next k

          'Debug.Print p; m; x; n; k; stF & " : " & Criteria(i)
        
        p = 0
       
       End If
 _

        p = 0
        x = 0
        k = 0
        m = 0
        n = 0
Next i

p = 0
m = 0
x = 0
y = 0
n = 0

Next counter
    counter = 0
    finish = Timer
    
    MsgBox finish - start
    Call ExactMatchWhole
    
End Sub

Sub ExactMatchWhole()
Application.ScreenUpdating = False

Dim Criteria() As Variant
Dim counter As Long
Dim i, j, n, m As Long
Dim st, stC As String
Dim s As Variant
Dim x As Variant
Dim r, rCrit As Range
Dim start, finish
start = Timer

Set rCrit = ActiveWorkbook.Sheets("sheet2").Range("Criteria2")

ReDim Criteria(0 To 1)
        For i = 0 To rCrit.Cells.Count - 1
        ReDim Preserve Criteria(0 To i)
        Criteria(i) = rCrit.Cells(i + 1).Value
        Next i

Set r = ActiveWorkbook.Sheets("ICMS Search").Range("LookIn")

For counter = 1 To r.Cells.Count
st = r.Cells(counter).Value

For i = 0 To UBound(Criteria)
If Not Criteria(i) = "" Then
On Error Resume Next
m = InStr(1, st, Criteria(i), vbBinaryCompare)
    If m > 0 Then

        r.Cells(counter).Characters(m, Len(Criteria(i))).Font.color = vbRed
    End If
End If
Next i
        
    Next counter
    
    finish = Timer
    
    MsgBox finish - start
End Sub

