Attribute VB_Name = "Module1"
Public Function ResizeLW(iFrm As Form, iLw As ListView, Optional bUseEmptyAreas As Boolean = True)

On Error Resume Next
iFrm.Font.Size = iLw.Font.Size
iFrm.Font.Name = iLw.Font.Name

Dim ColumnCount As Integer
Dim maxLen() As Long, totalLen As Long, tmpLen As Long
Dim X As Integer, Y As Integer
Dim tmpTxt As String
ColumnCount = iLw.ColumnHeaders.Count
ReDim maxLen(1 To ColumnCount)



For X = 1 To ColumnCount
    For Y = 1 To iLw.ListItems.Count
        If X = 1 Then
            tmpTxt = iLw.ListItems(Y).Text
        Else
            tmpTxt = iLw.ListItems(Y).ListSubItems(X - 1).Text
        End If
        tmpLen = iFrm.TextWidth(tmpTxt) + 200
        If tmpLen > maxLen(X) Then: maxLen(X) = tmpLen
    Next Y
    iLw.ColumnHeaders(X).Width = maxLen(X)
    totalLen = totalLen + maxLen(X)
Next X

If bUseEmptyAreas = True And (totalLen < iLw.Width - 200) Then
    Dim ShareVal As Long
    ShareVal = Int((iLw.Width - totalLen - 200) / ColumnCount)
    For X = 1 To ColumnCount
        iLw.ColumnHeaders(X).Width = iLw.ColumnHeaders(X).Width + ShareVal
    Next
End If

End Function

