Sub 一键修改成标准房号格式()
    Dim cell As Range
    Dim roomNumbers() As String
    Dim roomNumber As Variant
    Dim standardizedRoomNumbers As String
    
    On Error Resume Next ' Enable error handling
    
    For Each cell In Selection
        roomNumbers = Split(cell.Value, Chr(10))
        standardizedRoomNumbers = ""
        
        For Each roomNumber In roomNumbers
            standardizedRoomNumber = StandardizeRoomNumber(CStr(roomNumber))
            standardizedRoomNumbers = standardizedRoomNumbers & standardizedRoomNumber & Chr(10)
        Next roomNumber
        
        If Err.Number = 0 Then
            cell.Value = "'" & Left(standardizedRoomNumbers, Len(standardizedRoomNumbers) - 1) ' Add an apostrophe before the value to force text format
        Else
            Err.Clear ' Clear the error if any
        End If
    Next cell
    
    On Error GoTo 0 ' Disable error handling
End Sub

Function StandardizeRoomNumber(roomNumber As String) As String
    Dim parts() As String
    roomNumber = Replace(roomNumber, "#", "") ' Remove any '#' characters
    roomNumber = Replace(roomNumber, " ", "-") ' Replace any spaces with hyphens
    parts = Split(roomNumber, "-")
    
    If UBound(parts) = 3 Then
        ' Remove the unnecessary part and reconstruct the room number
        roomNumber = parts(0) & "-" & parts(1) & "-" & parts(3)
        parts = Split(roomNumber, "-")
    End If
    
    If UBound(parts) = 2 Then
        floorNumber = CStr(CInt(Left(parts(2), Len(parts(2)) - 2)))
        roomNumber = Right(parts(2), 2)
        StandardizeRoomNumber = parts(0) & "-" & parts(1) & "-" & floorNumber & roomNumber
    Else
        StandardizeRoomNumber = roomNumber
    End If
End Function