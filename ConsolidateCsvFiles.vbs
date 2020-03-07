option explicit

Const vbext_ct_ClassModule = 2
Const vbext_ct_Document = 100
Const vbext_ct_MSForm = 3
Const vbext_ct_StdModule = 1

Main

Sub Main
    'Read from config.txt
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

    Dim sCurPath
    sCurPath = fso.GetAbsolutePathName(".") 
    
    Wscript.Echo sCurPath
        
    Dim dicIni : Set dicIni = ReadIniFile(".\config.ini")
    Dim sSec, sKV
    For Each sSec In dicIni.Keys()
        WScript.Echo "---ini section:", sSec
        For Each sKV In dicIni(sSec).Keys()
            WScript.Echo " ", sKV, "=>", dicIni(sSec)(sKV)
        Next
    Next

    dim strPath, strOutputPath
    strPath = dicIni("csv")("path")
    strOutputPath = dicIni("csv")("output")
    if strOutputPath="" then
        strOutputPath = strPath
    end if

    WScript.Echo "strPath:", strPath, " => strOutputPath:", strOutputPath
    
    If Not fso.FolderExists(strPath) OR Not fso.FolderExists(strOutputPath) Then
        Wscript.Echo "Path or output not exist, please check strPath:", strPath, " => strOutputPath:", strOutputPath
	    Wscript.Quit
    end if

    dim today, strToday 
    today = GetFormattedDate()
    strToday = replace(today,"/","")
    Wscript.Echo " today:",today, "=>",strToday

    dim outPutFileName, outPutFileNameAmount
    outPutFileName = strOutputPath & "\OutputFile_"& strToday &".csv"
    outPutFileNameAmount = strOutputPath & "\OutputFileAmount_"& strToday &".csv"
    WScript.Echo " outPutFileName:",outPutFileName

    ''Find all csv file
    WScript.Echo "Working in directory: " & strPath
    
    dim totalAmount
    totalAmount = 0.0
     
    Dim objFolder, outputFile, outputFileAmount, objFile, inputFile
    Set objFolder = fso.GetFolder(strPath)
        'Set outputFile = fso.OpenTextFile(outPutFileName, 2, True) 'write/replace - don't append - create
        Set outputFile =  fso.OpenTextFile(outPutFileName, 2, False, -1)
        Dim strHeader
        For Each objFile in objFolder.Files
            WScript.Echo "Processing file: " & objFile.Name
            If LCase(Right(objFile.Name, 4)) = LCase(".csv") Then 'only for .CSV files
                WScript.Echo "Processing file for headers: " & objFile.Name
                Set inputFile = fso.OpenTextFile(strPath & "\" & objFile.Name, 1) 'reading
                    strHeader = inputFile.ReadLine
                    WScript.Echo strHeader

                    if Right(strHeader, 1) <> "," then
                        strHeader = strHeader & ","
                    end if

                    strHeader=strHeader & "ProcessDate"
                    WScript.Echo strHeader
                
                inputFile.Close
                Set inputFile = Nothing
                Exit For
            End If
        Next

        'WScript.Echo "Split for,", Join(Split(strHeader, ",", -1, 1),",")       
        outputFile.WriteLine strHeader 'let's write our first line - the headers we have
        WScript.Echo "..Finished processing the headers - wrote to file: " & OutPutFileName
        
        dim dataLine,arrDataLine
        For Each objFile in objFolder.Files
            If LCase(Right(objFile.Name, 4)) = LCase(".csv") Then 'only for .CSV files
                WScript.Echo "Processing file for data: " & objFile.Name
                Set inputFile = fso.OpenTextFile(strPath & "\" & objFile.Name, 1) 'reading
                    dataLine = inputFile.ReadLine
                    WScript.Echo "..ignore header line" & dataLine

                    Do While Not inputFile.AtEndOfStream
                        dataLine = inputFile.ReadLine
                        WScript.Echo dataLine

                        'handle total amount, assume the 2nd field if amount CDbl(x)
                        arrDataLine = Split(dataLine, ",", -1, 1)
                        if IsNumeric(arrDataLine(1)) AND Not Instr(1, arrDataLine(0), "合", 1) > 0 then
                            WScript.Echo "is number, do sum", arrDataLine(0),"=>", arrDataLine(1)
                            totalAmount = totalAmount + CDbl(arrDataLine(1))
                        else
                            WScript.Echo "not number, ignore", arrDataLine(0),"=>", arrDataLine(1)
                        end if

                        if Right(dataLine, 1) <> "," then
                            dataLine = dataLine & ","
                        end if

                        'dataLine=dataLine & """" & today & """"
                        dataLine=dataLine & today
                        'WScript.Echo dataLine  
                        outputFile.WriteLine dataLine
                    Loop
                inputFile.Close
                Set inputFile = Nothing
            End If

            'WScript.Echo "exit for check only"
            'Exit For
        Next

        outputFile.Close
        Set outputFile = Nothing
        WScript.Echo "..Finished - wrote to file: " & OutPutFileName
        
        Set outputFileAmount = fso.OpenTextFile(outPutFileNameAmount, 2, True)
        outputFileAmount.Write "Total Amount, "
        outputFileAmount.Write totalAmount
        outputFileAmount.Close
        Set outputFileAmount = Nothing
        
        WScript.Echo "..Finished - wrote total amount to file: " & outPutFileNameAmount

    Set objFolder = Nothing
    Set fso = Nothing

    WScript.Quit
End Sub

Function ReadIniFile(sFSpec)
    Dim goFS : Set goFS = CreateObject("Scripting.FileSystemObject")
    Dim dicTmp : Set dicTmp = CreateObject("Scripting.Dictionary")
    Dim tsIn   : Set tsIn   = goFS.OpenTextFile(sFSpec)
    Dim sLine, sSec, aKV
    Do Until tsIn.AtEndOfStream
        sLine = Trim(tsIn.ReadLine())
        If "[" = Left(sLine, 1) Then
            sSec = Mid(sLine, 2, Len(sLine) - 2)
            Set dicTmp(sSEc) = CreateObject("Scripting.Dictionary")
        Else
            If "" <> sLine Then
            aKV = Split(sLine, "=")
            If 1 = UBound(aKV) Then
                dicTmp(sSec)(Trim(aKV(0))) = Trim(aKV(1))
            End If
            End If
        End If
    Loop
    tsIn.Close
    Set goFS = Nothing 
    Set ReadIniFile = dicTmp
End Function

Function GetFormattedDate
    dim strDate, strDay, strMonth, strYear
    strDate = CDate(Date())
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