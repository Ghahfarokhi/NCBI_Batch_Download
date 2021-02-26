Attribute VB_Name = "C_Cleaning"
Option Explicit

Function Clean_Log()
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error GoTo Error_Handler
    Err.Number = 0
    '********************************************
    '********** HERE WE GO CLEANING! ************
    '********************************************
    CheckWorksheet ("Log")
    Sheets("Log").Activate
    ActiveWindow.DisplayGridlines = False
    
    With Sheets("Log").Range("A:A")
        .ClearContents
        .ColumnWidth = 150
        .ClearFormats
    End With
    
    Sheets("Log").Range("A1") = "Events log:"
    Sheets("Log").Range("A2") = "Date and Time\Procedure\info or error description:"
    Sheets("Log").Range("A2").Style = "Accent1"
    Sheets("Log").Range("A2").Font.Bold = True
    
    Event_Number = 0
    Sheets("RefSeq").Activate
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        On Error Resume Next
        User_Notification = "Clean_Log\Error Description: " & Err.Description
        Call Print_Log(0, User_Notification, "Bad")
        MsgBox "Something went wrong! Please check the Log worksheet for details!", vbExclamation, Tool_Name
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Function

Function Clean_Main()
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error GoTo Error_Handler
    Err.Number = 0
    '********************************************
    '******* CLEAN AND REFRESH THE FORMAT *******
    '********************************************
    CheckWorksheet ("Log")
    Sheets("RefSeq").Activate
    
    With Sheets("RefSeq").Range("A:AAA")
        .ClearContents
        .ClearFormats
    End With
    
    Sheets("RefSeq").Range(Range("Assembly"), Range("Strand").Offset(1000, 0)).Style = "Note"
    Sheets("RefSeq").Range(Range("File_Name"), Range("Annotation_Type").Offset(1000, 0)).Style = "Input"
    Sheets("RefSeq").Range(Range("File_Address"), Range("Comments").Offset(1000, 0)).Style = "Good"
    
    With Sheets("RefSeq").Range(Range("Assembly"), Range("Comments").Offset(1000, 0))
        .NumberFormat = "@"
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).ThemeColor = 1
        .Borders(xlInsideVertical).TintAndShade = -0.35
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).ThemeColor = 1
        .Borders(xlInsideHorizontal).TintAndShade = -0.35
        .Borders(xlInsideHorizontal).Weight = xlThin
    End With
    
    With Sheets("RefSeq").Range(Range("Assembly"), Range("Comments"))
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Borders(xlEdgeTop).ColorIndex = 0
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).ColorIndex = 0
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    With Sheets("RefSeq").Range(Range("Coordinate_Start"), Range("Coordinate_Stop").Offset(1000, 0))
        .NumberFormat = "General"
    End With
    
    Sheets("RefSeq").Range("Assembly").Offset(-1, 0) = "Required*"
    Sheets("RefSeq").Range("Assembly") = "Assembly*"
    Sheets("RefSeq").Range("Chromosome") = "Chromosome*"
    Sheets("RefSeq").Range("Coordinate_Start") = "Start*"
    Sheets("RefSeq").Range("Coordinate_Stop") = "End*"
    Sheets("RefSeq").Range("Strand") = "Strand*"
    
    Sheets("RefSeq").Range("File_Name").Offset(-1, 0) = "Optional:"
    Sheets("RefSeq").Range("File_Name") = "File Name"
    Sheets("RefSeq").Range("Annotation_Seq") = "Sequence to Annotate"
    Sheets("RefSeq").Range("Annotation_Name") = "Annotation Name"
    Sheets("RefSeq").Range("Annotation_Type") = "Annotation Type"
    
    Sheets("RefSeq").Range("File_Address").Offset(-1, 0) = "Output:"
    Sheets("RefSeq").Range("File_Address") = "File Address"
    Sheets("RefSeq").Range("Sequence") = "Sequence"
    Sheets("RefSeq").Range("Comments") = "Comments"
    
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        Sheets("Log").Range("A3").Offset(Event_Number, 0) = Now & "\Clean_Main_Worksheet\Error Description: " & Err.Description
        Sheets("Log").Range("A3").Offset(Event_Number, 0).Style = "Bad"
        Event_Number = Event_Number + 1
        MsgBox "Something went wrong! Please check the Log worksheet for details!", vbExclamation, Tool_Name
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Function

Function Format_Main_Wsh()
    
    On Error Resume Next
        
    Sheets("RefSeq").Range(Range("Assembly").Offset(1, 0), Range("Strand").Offset(1000, 0)).Style = "Note"
    Sheets("RefSeq").Range(Range("File_Name").Offset(1, 0), Range("Annotation_Type").Offset(1000, 0)).Style = "Input"
    Sheets("RefSeq").Range(Range("File_Address").Offset(1, 0), Range("Comments").Offset(1000, 0)).Style = "Good"
    Sheets("RefSeq").Range(Range("File_Address").Offset(1, 0), Range("Comments").Offset(1000, 0)).ClearContents

End Function


Function CheckWorksheet(Wsh As String)
    
    Dim ws As Worksheet
    Err.Number = 0
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Wsh)
    
    If Not Err.Number = 0 Then
        Sheets.Add.Name = Wsh
    End If
    
End Function

