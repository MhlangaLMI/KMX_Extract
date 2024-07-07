Attribute VB_Name = "Module11"
Option Explicit

Sub ExtractPoints()
    Dim kmzFilePath As String
    Dim xlFilePath As String
    Dim pythonExePath As String
    Dim scriptPath As String
    Dim coordinates As Variant
    Dim result As String
    Dim ws As Worksheet
    Dim startCell As Range
    Dim row As Long
    Dim col As Long
    Dim folderPath As String
    Dim dateTimeSelected As String
    Dim startTime As Double
    Dim endTime As Double
    Dim elapsedTime As Double
    Dim i As Long
    
    ' Disable screen updating to improve performance
    Application.ScreenUpdating = False
    
    ' Record the start time
    startTime = Timer
    
    ' Get the path of the current Excel file
    xlFilePath = ThisWorkbook.Path

    ' Open file dialog to select KMZ file
    kmzFilePath = Application.GetOpenFilename("KMZ Files (*.kmz), *.kmz", , "Select KMZ File")

    If kmzFilePath = "False" Then
        MsgBox "No file selected.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Extract folder path from selected KMZ file path
    folderPath = kmzFilePath
    
    ' Get the current date and time
    dateTimeSelected = Now

    ' Report the full path of the KMZ file and date/time of selection in cells C1 and C2
    Set ws = ActiveSheet
    ws.Range("C1").Value = folderPath
    ws.Range("C2").Value = dateTimeSelected

    ' Specify the path to the Python executable and the script
    pythonExePath = "python" ' Ensure Python is added to your PATH
    scriptPath = xlFilePath & "\extract_points.py"

    ' Run the Python script and get the result
    result = ShellAndWait(pythonExePath & " """ & scriptPath & """ """ & kmzFilePath & """", vbHide)

    ' Split the result into coordinates
    coordinates = Split(result, vbCrLf)

    ' Ensure coordinates array is not empty
    If UBound(coordinates) < 0 Then
        MsgBox "No coordinates found.", vbExclamation
        Application.ScreenUpdating = True
        Set ws = Nothing
        Exit Sub
    End If

    ' Start reporting coordinates from cell A5
    Set startCell = ws.Range("A5")
    row = startCell.row
    col = startCell.Column

    ' Write coordinates to the starting cell (A5) and corresponding data to other columns
    For i = LBound(coordinates) To UBound(coordinates)
        If coordinates(i) <> "" Then
            Dim lat As Double
            Dim lon As Double
            Dim elevation As Double
            Dim dms_lat As String
            Dim dms_lon As String
            Dim distance As Double
            Dim cumulative_distance As Double
            Dim name As String
            Dim coordParts As Variant
            coordParts = Split(coordinates(i), ",")
            
            If UBound(coordParts) = 7 Then
                On Error Resume Next
                lat = CDbl(coordParts(0))
                lon = CDbl(coordParts(1))
                elevation = CDbl(coordParts(2))
                dms_lat = coordParts(3)
                dms_lon = coordParts(4)
                distance = CDbl(coordParts(5))
                cumulative_distance = CDbl(coordParts(6))
                name = coordParts(7)
                On Error GoTo 0
                
                If Not IsEmpty(lat) And Not IsEmpty(lon) And Not IsEmpty(elevation) Then
                    ws.Cells(row, col).Value = name
                    ws.Cells(row, col + 1).Value = lat
                    ws.Cells(row, col + 2).Value = lon
                    ws.Cells(row, col + 3).Value = dms_lat
                    ws.Cells(row, col + 4).Value = dms_lat
                    ws.Cells(row, col + 5).Value = elevation
                    
                    row = row + 1
                End If
            Else
                MsgBox "Unexpected coordinate format: " & coordinates(i), vbExclamation
            End If
        End If
    Next i

    ' Record the end time and calculate elapsed time
    endTime = Timer
    elapsedTime = endTime - startTime
    
    ' Report elapsed time in HH:MM:SS format in cell E2
    ws.Range("E2").Value = Format$(elapsedTime / 86400, "HH:MM:SS")
    
    ' Clean up
    Set ws = Nothing
    Set startCell = Nothing

    ' Re-enable screen updating
    Application.ScreenUpdating = True
End Sub

Function ShellAndWait(ByVal cmd As String, ByVal windowStyle As VbAppWinStyle) As String
    Dim wsh As Object
    Dim exec As Object
    Dim output As String
    Set wsh = CreateObject("WScript.Shell")
    Set exec = wsh.exec(cmd)
    Do While exec.Status = 0
        DoEvents
        If Not exec.StdOut.AtEndOfStream Then
            output = output & exec.StdOut.ReadAll
        End If
    Loop
    ShellAndWait = output
    
    ' Clean up
    Set exec = Nothing
    Set wsh = Nothing
End Function

