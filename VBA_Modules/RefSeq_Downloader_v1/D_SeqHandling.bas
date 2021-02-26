Attribute VB_Name = "D_SeqHandling"
Option Explicit

Function Annotator(Batch As Long, GenBank_File_Path As String, sgRNA As String, AnnotationName As String, AnnotationType As String, Locus_Name As String) As Boolean
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.DisplayAlerts = False
    On Error GoTo Error_Handler
    Err.Number = 0
    '********************************************
    '*******             MAIN             *******
    '********************************************
    Annotator = False
    
    Dim Seq As String, RevCompSeq As String, DataLine As String, i As Integer
    Dim ORIGIN_Found As Boolean
    Dim CRISPR_Features_Added As Boolean
    Dim sgRNA_Strand As String, sgRNA_Feature As String
    
    Dim Annotated_File_Path As String, Temp_File_Path As String
    
    If AnnotationName = "" Then AnnotationName = "Annotation_Name_" & Str(Batch)
    If AnnotationType = "" Then AnnotationType = "Misc_Annotation"
    If Locus_Name = "" Then Locus_Name = "Locus_Name_" & Str(Batch)
    
    AnnotationName = Replace(AnnotationName, " ", "")
    AnnotationType = Replace(AnnotationType, " ", "")
    Locus_Name = Replace(Locus_Name, " ", "")
    
    Annotated_File_Path = GenBank_File_Path & "_Annotated.gb"
    Temp_File_Path = GenBank_File_Path & "temp.txt"
    
    ORIGIN_Found = False
    CRISPR_Features_Added = False
    
    Open GenBank_File_Path For Input As #1
    Open Temp_File_Path For Output As #2
    
    While Not EOF(1)
        Line Input #1, DataLine
        DataLine = Replace(DataLine, Chr(10), vbCrLf)
        Print #2, DataLine
    Wend
    
    Close #1
    Close #2
    
    Open Temp_File_Path For Input As #1
    
    While Not EOF(1)
        Line Input #1, DataLine
        If ORIGIN_Found = False And (DataLine = "ORIGIN" Or InStr(1, DataLine, "ORIGIN ") > 0) Then
            ORIGIN_Found = True
            Seq = Mid(DataLine, InStr(1, DataLine, "ORIGIN ") + 7, Len(DataLine))
            GoTo Next_Loop
        ElseIf ORIGIN_Found = True Then
            Seq = Seq + DataLine
        End If
Next_Loop:
    Wend
    Close #1
    
    For i = 0 To 9
        Seq = Replace(Seq, i, "")
    Next i
    
    Seq = Replace(Seq, Chr(10), "")
    Seq = Replace(Seq, " ", "")
    Seq = Replace(Seq, "/", "")
    
    Seq = UCase(Seq)
    RevCompSeq = RevComp(Seq)
    
    If InStr(1, Seq, sgRNA) > 0 Then
        sgRNA_Strand = "Fwd"
    ElseIf InStr(1, RevCompSeq, sgRNA) > 0 Then
        sgRNA_Strand = "Rev"
    Else
        GoTo Exit_sub
    End If
    
    
    If sgRNA_Strand = "Fwd" Then
        sgRNA_Feature = Str(InStr(1, Seq, sgRNA)) & ".." & Str(InStr(1, Seq, sgRNA) - 1 + Len(sgRNA))
        sgRNA_Feature = "     " & AnnotationType & " " & Replace(sgRNA_Feature, " ", "")
    ElseIf sgRNA_Strand = "Rev" Then
        sgRNA_Feature = Str(InStr(1, Seq, RevComp(sgRNA))) & ".." & Str(InStr(1, Seq, RevComp(sgRNA)) - 1 + Len(sgRNA))
        sgRNA_Feature = "     " & AnnotationType & " complement(" & Replace(sgRNA_Feature, " ", "") & ")"
    End If
    
    AnnotationName = "     /label=" & AnnotationName
    
    Open Temp_File_Path For Input As #1
    Open Annotated_File_Path For Output As #2
    
    While Not EOF(1)
        Line Input #1, DataLine
        If Left(DataLine, 5) = "LOCUS" Then
            DataLine = Left(DataLine, 12) + Locus_Name + Right(DataLine, Len(DataLine) - 25)
        End If
        If CRISPR_Features_Added = False Then
            If Not InStr(1, DataLine, "FEATURES") > 0 Then
                Print #2, DataLine
            Else
                Print #2, DataLine
                    If Not sgRNA_Strand = "Not found!" Then
                        Print #2, sgRNA_Feature
                        Print #2, AnnotationName
                    End If
                CRISPR_Features_Added = True
            End If
        Else
            Print #2, DataLine
        End If
    Wend
    
    Close #1
    Close #2
    
    Call Print_Log(Batch, "Annotation succeded!", "Good")
    
Exit_sub:
    Kill (Temp_File_Path)

    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        On Error Resume Next
        Annotator = False
        User_Notification = "Annotator\Error Description: " & Err.Description
        Call Print_Log(Batch, User_Notification, "Bad")
        Sheets("RefSeq").Range("Comments").Offset(Batch, 0).Style = "Bad"
        Err.Number = 0
    Else
        Annotator = True
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Application.DisplayAlerts = True
    
End Function


Function Seq_Extractor(Batch As Long, GenBank_File_Path As String) As String
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.DisplayAlerts = False
    On Error GoTo Error_Handler
    Err.Number = 0
    '********************************************
    '*******             MAIN             *******
    '********************************************
    Dim Seq As String, DataLine As String, i As Integer
    Dim ORIGIN_Found As Boolean
    
    Dim Temp_File_Path As String
    
    Temp_File_Path = GenBank_File_Path & "temp.txt"
    
    ORIGIN_Found = False
    
    Open GenBank_File_Path For Input As #1
    Open Temp_File_Path For Output As #2
    
    While Not EOF(1)
        Line Input #1, DataLine
        DataLine = Replace(DataLine, Chr(10), vbCrLf)
        Print #2, DataLine
    Wend
    
    Close #1
    Close #2
    
    Open Temp_File_Path For Input As #1
    
    While Not EOF(1)
        Line Input #1, DataLine
        If ORIGIN_Found = False And (DataLine = "ORIGIN" Or InStr(1, DataLine, "ORIGIN ") > 0) Then
            ORIGIN_Found = True
            Seq = Mid(DataLine, InStr(1, DataLine, "ORIGIN ") + 7, Len(DataLine))
            GoTo Next_Loop
        ElseIf ORIGIN_Found = True Then
            Seq = Seq + DataLine
        End If
Next_Loop:
    Wend
    Close #1
    
    For i = 0 To 9
        Seq = Replace(Seq, i, "")
    Next i
    
    Seq = Replace(Seq, Chr(10), "")
    Seq = Replace(Seq, " ", "")
    Seq = Replace(Seq, "/", "")
    
    Kill (Temp_File_Path)
    
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        'MsgBox Err.Description
        On Error Resume Next
        Seq_Extractor = "Seq_Extractor failed!"
        User_Notification = "Seq_Extractor\Error Description: " & Err.Description
        Call Print_Log(Batch, User_Notification, "Bad")
        Sheets("RefSeq").Range("Comments").Offset(Batch, 0).Style = "Bad"
        Err.Number = 0
    Else
        Seq_Extractor = UCase(Seq)
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Application.DisplayAlerts = True
    
End Function





