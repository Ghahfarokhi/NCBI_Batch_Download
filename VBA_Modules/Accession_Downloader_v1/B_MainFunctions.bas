Attribute VB_Name = "B_MainFunctions"
Option Explicit

Function Print_Log(i As Long, Msg As String, Style_Format As String)
    
    On Error Resume Next
    
    Sheets("Log").Range("A3").Offset(Event_Number, 0) = Now & "\" & Str(i) & " \" & Msg
    Sheets("Log").Range("A3").Offset(Event_Number, 0).Style = Style_Format
    Event_Number = Event_Number + 1
    
End Function


Function Count_Records() As Boolean
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.ScreenUpdating = False
    On Error GoTo Error_Handler
    Err.Number = 0
    '********************************************
    '*******             MAIN             *******
    '********************************************
    Total_Records = 0
    
    Dim i As Long
    
    i = Range("Accession").End(xlDown).Row - Sheets("Accession").Range("Accession").Row
    
    If i < 0 Then i = 0
    If i > 100000 Then i = 0
    If i > 1000 Then i = 1000
    
    If i >= 1 Then
        Total_Records = i
        Call Format_Main_Wsh
        Count_Records = True
        User_Notification = "Total number of records: " & Str(i)
        Call Print_Log(0, User_Notification, "Good")
    End If
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        On Error Resume Next
        User_Notification = "Count_Records\Error Description: " & Err.Description
        Call Print_Log(0, User_Notification, "Bad")
        Count_Records = False
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Application.DisplayAlerts = True
    
End Function

Function Check_Inputs(i As Long) As Boolean
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.ScreenUpdating = False
    On Error GoTo Error_Handler
    Err.Number = 0
    '********************************************
    '******           MAIN CODE!           ******
    '********************************************
    Check_Inputs = False
    
    Accession = Sheets("Accession").Range("Accession").Offset(i, 0)
    
    If Len(Accession) < 3 Then
        User_Notification = "Please check the provided Accession."
        Call Print_Log(i, User_Notification, "Bad")
        Sheets("Accession").Range("Comments").Offset(i, 0) = User_Notification
        Sheets("Accession").Range("Comments").Offset(i, 0).Style = "Bad"
        GoTo Error_Handler
    End If
    
    Databank = UCase(Sheets("Accession").Range("Databank").Offset(i, 0))
    
    If Not (Databank = "NUCLEOTIDE" Or Databank = "NUCCORE" Or Databank = "PROTEIN") Then
        User_Notification = "Unvalid Databank! Valid values: Nucleotide/Nuccore/Protein."
        Call Print_Log(i, User_Notification, "Bad")
        Sheets("Accession").Range("Comments").Offset(i, 0) = User_Notification
        Sheets("Accession").Range("Comments").Offset(i, 0).Style = "Bad"
        GoTo Error_Handler
    End If
    
    Position_Start = CLng(Sheets("Accession").Range("Coordinate_Start").Offset(i, 0))
    Position_End = CLng(Sheets("Accession").Range("Coordinate_Stop").Offset(i, 0))
    
    If Position_Start > 0 And Position_End > 0 Then
    
        If (Position_End - Position_Start) <= 0 Then
            User_Notification = "Unvalid coordinates! RefSeq Length <=0"
            Call Print_Log(i, User_Notification, "Bad")
            Sheets("Accession").Range("Comments").Offset(i, 0) = User_Notification
            Sheets("Accession").Range("Comments").Offset(i, 0).Style = "Bad"
            GoTo Error_Handler
        End If
        
        If (Position_End - Position_Start) > 300000 And ActiveSheet.Shapes("Both_Seq_GB").OLEFormat.Object.Value = xlOn Then
            User_Notification = "Unvalid coordinates! RefSeq Length > 300K. Only the first 300K will be downloaded."
            Call Print_Log(i, User_Notification, "Bad")
            Sheets("Accession").Range("Comments").Offset(i, 0) = User_Notification
            Sheets("Accession").Range("Comments").Offset(i, 0).Style = "Bad"
            Position_End = Position_Start + 300000
        End If
        
        If (Position_End - Position_Start) > 32767 And ActiveSheet.Shapes("Seq_Only").OLEFormat.Object.Value = xlOn Then
            User_Notification = "RefSeq Length > 32767 bp. Only the first 32767 bp will be fetched into the spreadsheet."
            Call Print_Log(i, User_Notification, "Neutral")
            Sheets("Accession").Range("Comments").Offset(i, 0) = User_Notification
            Sheets("Accession").Range("Comments").Offset(i, 0).Style = "Neutral"
            Position_End = Position_Start + 32767
        End If
        
        If (Position_End - Position_Start) > 32767 And ActiveSheet.Shapes("Both_Seq_GB").OLEFormat.Object.Value = xlOn Then
            User_Notification = "RefSeq Length > 32767 bp. Only the first 32767 bp will be fetched into the spreadsheet. However, the downloaded file will contain the full sequence."
            Call Print_Log(i, User_Notification, "Neutral")
            Sheets("Accession").Range("Comments").Offset(i, 0) = User_Notification
            Sheets("Accession").Range("Comments").Offset(i, 0).Style = "Neutral"
        End If
    
    Else
        
        Position_Start = 0
        Position_End = 0
        
    End If
    
    Ret_Type = UCase(Sheets("Accession").Range("Ret_Type").Offset(i, 0))
    
    
    If Ret_Type = "FASTA" Or Ret_Type <= "FASTA_CDS_NA" Or Ret_Type = "FASTA_CDS_AA" Then
        Extension = ".fasta"
    ElseIf Ret_Type = "GENBANK" Or Ret_Type = "GB" Then
        Ret_Type = "Genbank"
        Extension = ".gb"
    ElseIf Ret_Type = "GENPEPT" Or Ret_Type = "GP" Then
        Ret_Type = "Genpept"
        Extension = ".gp"
    Else
        Ret_Type = "FASTA"
        Extension = ".fasta"
    End If
    

    
    File_Name = Sheets("Accession").Range("File_Name").Offset(i, 0)
    File_Name = Replace(Replace(Replace(Replace(Replace(File_Name, "/", ""), "\", ""), ",", ""), ":", ""), ";", "")
    File_Name = Replace(Replace(Replace(Replace(Replace(File_Name, "*", ""), "<", ""), ">", ""), Chr(34), ""), "|", "")
    
    If File_Name = "" Then
        File_Name = Accession & "_" & Position_Start & "_" & Position_End
    ElseIf Len(File_Name) > 200 Then
        File_Name = Left(File_Name, 100) & "_" & Right(File_Name, 100)
    End If
    
    Check_Inputs = True
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        On Error Resume Next
        User_Notification = "Check_Inputs\Error Description: " & Err.Description
        Call Print_Log(i, User_Notification, "Bad")
        Sheets("Accession").Range("Comments").Offset(i, 0) = User_Notification
        Sheets("Accession").Range("Comments").Offset(i, 0).Style = "Bad"
        Check_Inputs = False
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Application.DisplayAlerts = True
    
End Function

Function Generate_URL(i As Long) As Boolean
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.ScreenUpdating = False
    On Error GoTo Error_Handler
    Err.Number = 0
    '********************************************
    '******           MAIN CODE!           ******
    '********************************************
    Generate_URL = False
    
    If Position_Start = 0 And Position_End = 0 Then
        GenBank_URL = "https://www.ncbi.nlm.nih.gov/sviewer/viewer.cgi?tool=portal&save=file&log$=seqview&db=" & Databank & _
                      "&report=" & Ret_Type & "&id=" & Accession & "&conwithfeat=on&withparts=on&hide-cdd=on"
    Else
        GenBank_URL = "https://www.ncbi.nlm.nih.gov/sviewer/viewer.cgi?tool=portal&save=file&log$=seqview&db=" & Databank & _
                      "&report=" & Ret_Type & "&id=" & Accession & "&from=" & Position_Start & "&to=" & Position_End & "&conwithfeat=on&withparts=on&hide-cdd=on"
    End If
    
    GenBank_URL = Replace(GenBank_URL, " ", "")
    
    Generate_URL = True

    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        On Error Resume Next
        User_Notification = "Generate_URL\Error Description: " & Err.Description
        Call Print_Log(i, User_Notification, "Bad")
        Sheets("Accession").Range("Comments").Offset(i, 0) = User_Notification
        Sheets("Accession").Range("Comments").Offset(i, 0).Style = "Bad"
        Generate_URL = False
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Application.DisplayAlerts = True
    
End Function

Function Download_File(i As Long, File_URL As String, Save_Address As String) As Boolean
    '********************************************
    '*************** ERROR POLICY ***************
    '********************************************
    Application.DisplayAlerts = False
    On Error Resume Next
    Err.Number = 0
    '********************************************
    '*******             MAIN             *******
    '********************************************
    Download_File = False
    
    Err.Number = 0
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", File_URL, False
    WinHttpReq.Send
    
    If WinHttpReq.Status = 200 Then
    
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile Save_Address, 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
        Download_File = True
    
    Else
        
        If Test_Connection(i, File_URL, Save_Address) = False Then
            Download_File = False
            Exit Function
        End If
        
    End If
    
    'Check if the saved_File exist here...
    Temp_Text = Dir(Save_Address)
    If Temp_Text = "" Then
        Download_File = False
        Exit Function
    Else
        Download_File = True
    End If
    
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        On Error Resume Next
        Download_File = False
        User_Notification = "Download_File\Error Description: " & Err.Description
        Call Print_Log(i, User_Notification, "Bad")
        Sheets("Accession").Range("Comments").Offset(i, 0).Style = "Bad"
        Err.Number = 0
    Else
        Download_File = True
    End If
    '********************************************
    '************* DEFAULT SETTINGS *************
    '********************************************
    Application.DisplayAlerts = True

End Function

Function Test_Connection(Batch As Long, Link As String, Optional Address_To_Save As String) As Boolean
    Test_Connection = False
    
    Err.Number = 0
    On Error Resume Next
    Dim WinHttpReq As Object
    
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", Link, False
    WinHttpReq.Send

    Temp_Counter = WinHttpReq.Status

    If Temp_Counter = 200 Then
        
        Call Print_Log(Batch, "Internet connection is Ok!", "Good")
        Test_Connection = True
        
            If Not Address_To_Save = "" Then
                Set oStream = CreateObject("ADODB.Stream")
                    oStream.Open
                    oStream.Type = 1
                    oStream.Write WinHttpReq.responseBody
                    oStream.SaveToFile Address_To_Save, 2 ' 1 = no overwrite, 2 = overwrite
                    oStream.Close
            End If
        
        Err.Number = 0
        Exit Function
    
    Else
        Call Print_Log(Batch, "Testing the internet connection failed!", "Bad")
        Call Connection_Aid(Batch, Temp_Counter)
    End If
    
    Err.Number = 0

End Function

Function Connection_Aid(Batch As Long, Req_Status As Long)
    
    On Error Resume Next
    
    If Req_Status = 200 Then Call Print_Log(Batch, "Internet connection status: OK.", "Good")
    If Req_Status = 100 Then Call Print_Log(Batch, "Internet connection status: Continue.", "Bad")
    If Req_Status = 101 Then Call Print_Log(Batch, "Internet connection status: Switching protocols.", "Bad")
    
    If Req_Status = 201 Then Call Print_Log(Batch, "Internet connection status: Created.", "Bad")
    If Req_Status = 202 Then Call Print_Log(Batch, "Internet connection status: Accepted.", "Bad")
    If Req_Status = 203 Then Call Print_Log(Batch, "Internet connection status: Non-Authoritative Information.", "Bad")
    If Req_Status = 204 Then Call Print_Log(Batch, "Internet connection status: No Content.", "Bad")
    If Req_Status = 205 Then Call Print_Log(Batch, "Internet connection status: Reset Content.", "Bad")
    If Req_Status = 206 Then Call Print_Log(Batch, "Internet connection status: Partial Content.", "Bad")
    
    If Req_Status = 300 Then Call Print_Log(Batch, "Internet connection status: Multiple Choices.", "Bad")
    If Req_Status = 301 Then Call Print_Log(Batch, "Internet connection status: Moved Permanently.", "Bad")
    If Req_Status = 302 Then Call Print_Log(Batch, "Internet connection status: Found.", "Bad")
    If Req_Status = 303 Then Call Print_Log(Batch, "Internet connection status: See Other.", "Bad")
    If Req_Status = 304 Then Call Print_Log(Batch, "Internet connection status: Not Modified.", "Bad")
    If Req_Status = 305 Then Call Print_Log(Batch, "Internet connection status: Use Proxy.", "Bad")
    If Req_Status = 307 Then Call Print_Log(Batch, "Internet connection status: Temporary Redirect.", "Bad")
    
    If Req_Status = 400 Then Call Print_Log(Batch, "Internet connection status: Bad Request.", "Bad")
    If Req_Status = 401 Then Call Print_Log(Batch, "Internet connection status: Unauthorized.", "Bad")
    If Req_Status = 402 Then Call Print_Log(Batch, "Internet connection status: Payment Required.", "Bad")
    If Req_Status = 403 Then Call Print_Log(Batch, "Internet connection status: Forbidden.", "Bad")
    If Req_Status = 404 Then Call Print_Log(Batch, "Internet connection status: Not Found.", "Bad")
    If Req_Status = 405 Then Call Print_Log(Batch, "Internet connection status: Method Not Allowed.", "Bad")
    If Req_Status = 406 Then Call Print_Log(Batch, "Internet connection status: Not Acceptable.", "Bad")
    If Req_Status = 407 Then Call Print_Log(Batch, "Internet connection status: Proxy Authentication Required.", "Bad")
    If Req_Status = 408 Then Call Print_Log(Batch, "Internet connection status: Request Timeout.", "Bad")
    If Req_Status = 409 Then Call Print_Log(Batch, "Internet connection status: Conflict.", "Bad")
    If Req_Status = 410 Then Call Print_Log(Batch, "Internet connection status: Gone.", "Bad")
    If Req_Status = 411 Then Call Print_Log(Batch, "Internet connection status: Length Required.", "Bad")
    If Req_Status = 412 Then Call Print_Log(Batch, "Internet connection status: Precondition Failed.", "Bad")
    If Req_Status = 413 Then Call Print_Log(Batch, "Internet connection status: Request Entity Too Large.", "Bad")
    If Req_Status = 414 Then Call Print_Log(Batch, "Internet connection status: Request-URI Too Long.", "Bad")
    If Req_Status = 415 Then Call Print_Log(Batch, "Internet connection status: Unsupported Media Type.", "Bad")
    If Req_Status = 416 Then Call Print_Log(Batch, "Internet connection status: Requested Range Not Suitable.", "Bad")
    If Req_Status = 417 Then Call Print_Log(Batch, "Internet connection status: Expectation Failed.", "Bad")
    
    If Req_Status = 500 Then Call Print_Log(Batch, "Internet connection status: Internal Server Error.", "Bad")
    If Req_Status = 501 Then Call Print_Log(Batch, "Internet connection status: Not Implemented.", "Bad")
    If Req_Status = 502 Then Call Print_Log(Batch, "Internet connection status: Bad Gateway.", "Bad")
    If Req_Status = 503 Then Call Print_Log(Batch, "Internet connection status: Service Unavailable.", "Bad")
    If Req_Status = 504 Then Call Print_Log(Batch, "Internet connection status: Gateway Timeout.", "Bad")
    If Req_Status = 505 Then Call Print_Log(Batch, "Internet connection status: HTTP Version Not Supported.", "Bad")
    
    Err.Number = 0
    
End Function

Function Defaulter()
    
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayFormulaBar = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Function

Function Check_Version()
    
    On Error GoTo Error_Handler
    
    Temp_Text = "https://raw.githubusercontent.com/Ghahfarokhi/NCBI_Batch_Download/master/version_Accession_Downloader.txt"
    Temp_File_Address = ActiveWorkbook.Path & "\" & "version_RefSeq_Downloader.txt"
    
    Dim Version_Line As String
    Dim New_Version As Long, Current_Version As Long
    
    Current_Version = CLng(Replace(Sheets("Info").Range("Version"), "Version: ", ""))
    
    If Download_File(0, Temp_Text, Temp_File_Address) = True Then
        
        Temp_Text = ""
        
        Open Temp_File_Address For Input As #1
        
        While Not EOF(1)
        
            Line Input #1, Version_Line
            Temp_Text = Temp_Text & Version_Line
            
        Wend
    
        Close #1
        
        Kill Temp_File_Address
        
        New_Version = CLng(Left(Temp_Text, InStr(1, Temp_Text, "/") - 1))
        
        If New_Version > Current_Version Then
            Call Print_Log(0, "A new version is available. Please download the updated version here: https://github.com/Ghahfarokhi/NCBI_Batch_Download", "Neutral")
        Else
            Call Print_Log(0, "RefSeq Downloader is up to date.", "Good")
        End If

    End If
   
    '********************************************
    '************** ERROR HANDLING **************
    '********************************************
Error_Handler:
    If Err.Number <> 0 Then
        On Error Resume Next
        User_Notification = Now & "\Check_Update\Error Description: " & Err.Description
        Call Print_Log(0, User_Notification, "Bad")
    End If
End Function

Function RevComp(RefSeq As String) As String
    
    On Error Resume Next
    
    Dim RCRefSeq As String
    
    RCRefSeq = Replace(UCase(RefSeq), " ", "")
    RCRefSeq = Replace(Replace(Replace(Replace(Replace(RCRefSeq, "A", "1"), "T", "2"), "C", "3"), "G", "4"), "U", "5")
    RCRefSeq = Replace(Replace(Replace(Replace(Replace(RCRefSeq, "1", "T"), "2", "A"), "3", "G"), "4", "C"), "5", "A")
    RCRefSeq = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(RCRefSeq, "Y", "1"), "R", "2"), "K", "3"), "M", "4"), "B", "5"), "V", "6"), "D", "7"), "H", "8")
    RCRefSeq = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(RCRefSeq, "1", "R"), "2", "Y"), "3", "M"), "4", "K"), "5", "V"), "6", "B"), "7", "H"), "8", "D")
    
    RevComp = StrReverse(RCRefSeq)

End Function

Function Translate1(DNA As String) As String
    
    On Error Resume Next
    
    Dim Codon As String, Translation As String, Length As Long
    Dim i As Long, j As Integer, Temp_Range As Range, Codon_Array() As Variant
    
    Set Temp_Range = Sheets("Info").Range(Range("Codon_Table").Offset(1, 0), Range("Codon_Table").Offset(64, 2))
    Codon_Array = Temp_Range
    
    DNA = Replace(Replace(Replace(UCase(DNA), "U", "T"), "-", ""), " ", "")
    
    Translation = ""
    
    Length = Len(DNA) \ 3

    For i = 1 To Length
    
        Codon = Mid(DNA, (i - 1) * 3 + 1, 3)
        
        For j = 1 To 64
            
            If Codon = Codon_Array(j, 1) Then
                Translation = Translation & Codon_Array(j, 2)
                If Codon_Array(j, 2) = "*" Then
                    GoTo Stop_Translation
                Else
                    GoTo Next_Codon
                End If
            End If
            
        Next j
        
Next_Codon:
    Next i
    
Stop_Translation:
    Translate1 = Translation
End Function

Function Translate3(DNA As String) As String
    
    On Error Resume Next
    
    Dim Codon As String, Translation As String, Length As Long
    Dim i As Long, j As Integer, Temp_Range As Range, Codon_Array() As Variant
    
    Set Temp_Range = Sheets("Info").Range(Range("Codon_Table").Offset(1, 0), Range("Codon_Table").Offset(64, 2))
    Codon_Array = Temp_Range
    
    DNA = Replace(Replace(Replace(UCase(DNA), "U", "T"), "-", ""), " ", "")
    
    Translation = ""
    
    Length = Len(DNA) \ 3

    For i = 1 To Length
    
        Codon = Mid(DNA, (i - 1) * 3 + 1, 3)
        
        For j = 1 To 64
            
            If Codon = Codon_Array(j, 1) Then
                Translation = Translation & Codon_Array(j, 3) & " "
                If Codon_Array(j, 2) = "Ter" Or Codon_Array(j, 3) = "Stp" Then
                    GoTo Stop_Translation
                Else
                    GoTo Next_Codon
                End If
            End If
            
        Next j
        
Next_Codon:
    Next i
    
Stop_Translation:
    Translate3 = Translation
End Function
