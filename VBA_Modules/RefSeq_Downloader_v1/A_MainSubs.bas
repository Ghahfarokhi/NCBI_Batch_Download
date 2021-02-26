Attribute VB_Name = "A_MainSubs"
    '#############################################
    '############# <  INFORMATION  > #############
    '#############################################
    '###                                       ###
    '###           RefSeq Downloader           ###
    '###              Version 1.0              ###
    '###              2020 May 02              ###
    '###                                       ###
    '###                                       ###
    '###               Author:                 ###
    '###        Amir Taheri Ghahfarokhi        ###
    '###                                       ###
    '###               Email:                  ###
    '###   Amir.Taheri.Ghahfarokhi@Gmail.com   ###
    '###                                       ###
    '###               GitHub                  ###
    '###    https://github.com/Ghahfarokhi/    ###
    '###                                       ###
    '###                                       ###
    '#############################################
    '#############################################
    '#############################################

'=======================================================
Option Explicit
'Declaring variables:
Public Const Tool_Name As String = "RefSeq Downloader v1.0"
Public Event_Number As Long
Public Total_Records As Long
Public User_Notification As String

Public Assembly As String
Public Chromosome As String
Public Position_Start As Double
Public Position_End As Double
Public Chr_Strand As String
Public File_Name As String
Public ValidChromosome As Boolean
Public Valid_Assembly As Boolean

Public Chr_ID_Array() As Variant

Public Gene_Length As Double
Public GeneID_Lib_Path As String
Public GenBank_URL As String

Public Sequence_To_Annotate As String
Public Extracted_Sequence As String
Public Annotation_Name As String
Public Annotation_Type As String
Public Locus_Name As String

Public WinHttpReq As Object
Public oStream As Object
Public Temp_File_Address As String
Public Temp_Counter As Long
Public Temp_Text As String

Public Start_Time As Double
Public Elapsed_Time As Double
'========================================================

Sub Download()
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.DisplayAlerts = False
    Application.StatusBar = False
    On Error GoTo Error_Handler
    Err.Number = 0
    Event_Number = 0
    Clean_Log
    Sheets("RefSeq").Activate
    '********************************************
    '*******             MAIN             *******
    '********************************************

    Dim i As Long, Temp_File_Address As String
    
    If Count_Records = False Then
        User_Notification = "Please complete the required fields!"
        MsgBox User_Notification, vbExclamation, Tool_Name
        GoTo Error_Handler
    End If
    
    If Load_Chr_ID_Array = False Then
        User_Notification = "Fatal error: Load_Chr_ID_Array failed! Please check the Log Worksheet."
        MsgBox User_Notification, vbExclamation, Tool_Name
        GoTo Error_Handler
    End If
    
    'Check the internet connection and NCBI response:
    Application.StatusBar = "Please wait: Checking the internet connection..."
    Temp_Text = "https://www.ncbi.nlm.nih.gov/sviewer/viewer.cgi?tool=portal&save=file&log$=seqview&db=nuccore&report=genbank&id=568815597&from=" & _
                Str(Int((500 - 100 + 1) * Rnd + 100)) & "&to=" & Str(Int((999 - 600 + 1) * Rnd + 600)) & "&"
    Temp_Text = Replace(Temp_Text, " ", "")
    If Test_Connection(0, Temp_Text) = False Then GoTo Error_Handler
    
    'Check version:
    Call Check_Version
    
    On Error GoTo Next_i
    Start_Time = Timer
    For i = 1 To Total_Records
        
        If i = 1 Then
            Application.StatusBar = "Downloading RefSeq " & Str(i) & "/" & Str(Total_Records)
        Else
            Elapsed_Time = (Timer - Start_Time) / (i - 1)
            Application.StatusBar = "Downloading RefSeq " & Str(i) & "/" & Str(Total_Records) & " , Remaining time: " & Format((Total_Records - i) * Elapsed_Time / 86400 + 10, "hh:mm:ss")
        End If
        
        DoEvents
        
        If Check_Inputs(i) = False Then
            User_Notification = "Checking input failed!"
            Call Print_Log(i, "Checking input failed!", "Bad")
            Sheets("RefSeq").Range("Comments").Offset(i, 0) = User_Notification
            Sheets("RefSeq").Range("Comments").Offset(i, 0).Style = "Bad"
            GoTo Next_i
        End If
                
        If Generate_URL(i) = False Then
            User_Notification = "Generating the URL failed!"
            Call Print_Log(i, User_Notification, "Bad")
            Sheets("RefSeq").Range("Comments").Offset(i, 0) = User_Notification
            Sheets("RefSeq").Range("Comments").Offset(i, 0).Style = "Bad"
            GoTo Next_i
        End If
        
        Temp_File_Address = ActiveWorkbook.Path & "\" & File_Name & ".gb"
        
        If Download_File(i, GenBank_URL, Temp_File_Address) = True Then
            Sheets("RefSeq").Range("Comments").Offset(i, 0) = "Download succeded!"
            Extracted_Sequence = Seq_Extractor(i, Temp_File_Address)
            Sheets("RefSeq").Range("Sequence").Offset(i, 0) = Extracted_Sequence
        Else
            If Test_Connection(i, GenBank_URL) = False Then
                Sheets("RefSeq").Range("Comments").Offset(i, 0) = "Download failed!"
                Sheets("RefSeq").Range("Comments").Offset(i, 0).Style = "Bad"
                GoTo Next_i
            End If
        End If
        
        If ActiveSheet.Shapes("Seq_Only").OLEFormat.Object.Value = xlOn Then
            
            Sheets("RefSeq").Range("File_Address").Offset(i, 0) = "Not applicable!"
            Kill (Temp_File_Address)
            
        ElseIf ActiveSheet.Shapes("Both_Seq_GB").OLEFormat.Object.Value = xlOn Then
            
            Sheets("RefSeq").Range("File_Address").Offset(i, 0) = Temp_File_Address
            Sheets("RefSeq").Range("File_Name").Offset(i, 0) = File_Name
            
            If Len(Extracted_Sequence) = (Position_End - Position_Start + 1) Then
                
                Sequence_To_Annotate = UCase(Sheets("RefSeq").Range("Annotation_Seq").Offset(i, 0))
                
                If Not Sequence_To_Annotate = "" Then
                
                    If InStr(1, Extracted_Sequence, Sequence_To_Annotate) > 0 Or InStr(1, RevComp(Extracted_Sequence), Sequence_To_Annotate) > 0 Then
                        Call Print_Log(i, "The Sequence_To_Annotate exists within the GenBank file!", "Good")
                        Annotation_Name = Sheets("RefSeq").Range("Annotation_Name").Offset(i, 0)
                        Annotation_Type = Sheets("RefSeq").Range("Annotation_Type").Offset(i, 0)
                        If Annotator(i, Temp_File_Address, Sequence_To_Annotate, Annotation_Name, Annotation_Type, File_Name) = True Then
                            Kill (Temp_File_Address)
                        Else
                            Call Print_Log(i, "Annotation failed!", "Bad")
                        End If
                    Else
                        Call Print_Log(i, "Couldn't find the Sequence_To_Annotate within the GenBank file!", "Bad")
                    End If

                End If

            End If
            
        End If
        
        Call Print_Log(i, "Procedure is complete!", "Good")

Next_i:

        If Err.Number <> 0 Then
            User_Notification = "Error" & Str(Err.Number) & ": " & Err.Description
            Sheets("RefSeq").Range("Comments").Offset(i, 0) = User_Notification
            Sheets("RefSeq").Range("Comments").Offset(i, 0).Style = "Bad"
            Call Print_Log(i, User_Notification, "Bad")
            Err.Number = 0
        End If
        
        If i Mod 20 = 0 Then ThisWorkbook.Save
    
    Next i
    
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        On Error Resume Next
        User_Notification = Now & "\Download\Error Description: " & Err.Description
        Call Print_Log(0, User_Notification, "Bad")
        MsgBox "Something went wrong! Please check the Log worksheet for details!", vbExclamation, Tool_Name
        Call Defaulter
        Exit Sub
    End If
    
    MsgBox "Done! Please check the Log Worksheet.", vbInformation, Tool_Name
    
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Call Defaulter
    
End Sub


Sub Clean_Workbook()
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error GoTo Error_Handler
    Err.Number = 0
    '********************************************
    '******* CALL MAIN CLEANING FUNCTIONS *******
    '********************************************
    
    Clean_Log
    Clean_Main
    
    MsgBox "Worksheets are cleaned!", vbInformation, Tool_Name
    
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        On Error Resume Next
        User_Notification = Now & "\Clean_Main_Worksheet\Error Description: " & Err.Description
        Call Print_Log(0, User_Notification, "Bad")
        MsgBox "Something went wrong! Please check the Log worksheet for details!", vbExclamation, Tool_Name
        Call Defaulter
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Call Defaulter
    
End Sub

