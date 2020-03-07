Attribute VB_Name = "modMain"
Sub btnDoConsolidate_Click()

    Application.ScreenUpdating = False
    Logging.log Application.ActiveWorkbook.FullName & "; consolidate called."
    
    'Dim sCurPath As String
    'sCurPath = Application.ActiveWorkbook.path
    'Logging.log sCurPath
    
    Range("B13") = ""
    Dim today, strToday
    today = GetFormattedDate()
    strToday = Replace(today, "/", "")
    'Logging.log " today: " & today & " => " & strToday
    
    Dim checkPath As String
    checkPath = Range("$B$2").Value
    
    If Right(checkPath, 1) <> "\" Then
        checkPath = checkPath & "\"
    End If
    
    Dim totalAmt As Double
    totalAmt = Range("$B$4").Value
    
    Dim strOutputPath As String
    strOutputPath = Range("$B$6").Value
    
    If strOutputPath = "" Then
        strOutputPath = checkPath
    End If
        
    If Right(strOutputPath, 1) <> "\" Then
        strOutputPath = strOutputPath & "\"
    End If
   
    Logging.log "checkPath: " & checkPath & ";totalAmt:" & Format(totalAmt, "#,##0.00") & ";strOutputPath:" & strOutputPath
        
    Dim outputFileName As String, outputFileNameAmount As String
    outputFileName = strOutputPath & "outputFile_" & strToday & ".csv"
    outputFileNameAmount = strOutputPath & "outputFileAmount_" & strToday & ".csv"
    
    FileExistDelete (outputFileName)
    FileExistDelete (outputFileNameAmount)
    
    Dim calTotalAmt As Double
    calTotalAmt = 0
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim outputFile As Object
    'Set outputFile = fso.CreateTextFile(outputFileName)
    Set outputFile = fso.OpenTextFile(outputFileName, ForAppending, True, TristateTrue)
        
        Dim firstFile As Boolean
        firstFile = True
           
        Dim csvFile As String
        Dim dataLine As String
        Dim arrDataLine As Variant
        Dim Cnt As Long
        
        csvFile = Dir(checkPath & "*.csv")
                
        Do While Len(csvFile) > 0
            Cnt = Cnt + 1
            Logging.log "Processing file: " & csvFile & ";first file:" & CStr(firstFile)
                        
            Open checkPath & csvFile For Input As #1
                Line Input #1, dataLine
                Debug.Print dataLine
                
                If firstFile Then
                    'process header.
                    Dim strHeader As String
                    Logging.log "Processing file for headers: " & csvFile

                    strHeader = dataLine
                    If Right(strHeader, 1) <> "," Then
                        strHeader = strHeader & ","
                    End If

                    strHeader = strHeader & "ProcessDate"
                    Logging.log strHeader

                    outputFile.WriteLine strHeader 'let's write our first line - the headers we have
                    Logging.log "..Finished processing the headers - wrote to file: " & outputFileName
                End If

                Do Until EOF(1)
                    Line Input #1, dataLine
                    Debug.Print dataLine

                    arrDataLine = Split(dataLine, ",")
                    'handle total amount, assume the 2nd field if amount CDbl(x)
                    If IsNumeric(arrDataLine(1)) And InStr(1, arrDataLine(0), "2019", 1) > 0 Then
                        Logging.log "is number, do sum" & arrDataLine(0) & "=>" & arrDataLine(1)
                        calTotalAmt = calTotalAmt + CDbl(arrDataLine(1))
                    Else
                        Logging.log "not number, ignore" & arrDataLine(0) & "=>" & arrDataLine(1)
                    End If

                    If Right(dataLine, 1) <> "," Then
                        dataLine = dataLine & ","
                    End If

                    'dataLine=dataLine & """" & today & """"
                    dataLine = dataLine & today
                    'Debug.Print dataLine
                    outputFile.WriteLine dataLine
                Loop
            Close #1
            
            firstFile = False
            csvFile = Dir()
        Loop
    
        outputFile.Close
        Set outputFile = Nothing
        
        Logging.log "..Finished - wrote to file: " & outputFileName
        
        Set outputFileAmount = fso.CreateTextFile(outputFileNameAmount)
            outputFileAmount.WriteLine "Total Amount, " & Format(totalAmt, "#,##0.00")
            outputFileAmount.WriteLine "Total Calculated Amount from files, " & Format(calTotalAmt, "#,##0.00")
            outputFileAmount.Close
        Set outputFileAmount = Nothing
        
        Logging.log "..Finished - wrote total amount to file: " & outputFileNameAmount
        
    Set fso = Nothing
    Application.ScreenUpdating = True
    
    If Cnt = 0 Then
        Logging.log "No CSV files were found..."
    End If
    
    MsgBox "Finished consolidate csv file to " & outputFileName
    
    Range("B13") = "Finished consolidate csv file to " & outputFileName
End Sub

Function GetFormattedDate()
    Dim strDate, strDay, strMonth, strYear
    strDate = CDate(Date)
    strDay = DatePart("d", strDate)
    strMonth = DatePart("m", strDate)
    strYear = DatePart("yyyy", strDate)
    If strDay < 10 Then
        strDay = "0" & strDay
    End If
    If strMonth < 10 Then
        strMonth = "0" & strMonth
    End If
    GetFormattedDate = strYear & "/" & strMonth & "/" & strDay
End Function

Sub FileExistDelete(filePath)
    If Dir(filePath) <> "" Then
        Kill filePath
    End If
End Sub

