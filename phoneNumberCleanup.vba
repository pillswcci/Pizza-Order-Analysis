Sub CleanUpPhoneNumbers()

    ' Macro to clean up phone numbers and extract extensions
    
    MsgBox ("Starting Process")
    
    ' find the number of rows that need to be processed
    
        Dim thisSheet As Worksheet
        Dim thisRange As Range
        Dim numberOfRows As Integer
        
        Set thisSheet = ThisWorkbook.ActiveSheet
        Set thisRange = thisSheet.UsedRange
        
        numberOfRows = thisRange.Rows.Count
        
        MsgBox ("Number of rows to process " & numberOfRows)
        
    ' adjust the column widths
        
        thisSheet.Columns("B").ColumnWidth = 20
        thisSheet.Columns("C").ColumnWidth = 20
        
    ' go through the rows and clean up the phone numbers and extract the extensions
    
        For iCount = 1 To numberOfRows
            If (iCount = 1) Then
            
                ' this is the first row, so we'll set some headers
                
                    Cells(iCount, 2).Value = "Cleaned Phone Number"
                    Cells(iCount, 3).Value = "Extracted Extension"
                    
            Else
                
                ' processing a phone number
                
                    ' set some variables
                    
                        Dim phoneNumberStr As String
                        Dim extensionStr As String
                        Dim rawPhoneNumberStr As String
                        
                        phoneNumberStr = ""
                        extensionStr = ""
                        rawPhoneNumberStr = Cells(iCount, 1).Value
                        
                    ' see if there's an extension
                        
                        rawPhoneNumberStr = LCase(rawPhoneNumberStr)
                        
                        Dim locationOfX As Integer
                        locationOfX = InStr(rawPhoneNumberStr, "x")
                        
                        If (locationOfX <> 0) Then
                            ' there was an extension
                                extensionStr = Mid( _
                                    rawPhoneNumberStr, _
                                    (locationOfX + 1), _
                                    Len(rawPhoneNumberStr) - locationOfX)
                                rawPhoneNumberStr = Replace( _
                                    rawPhoneNumberStr, _
                                    ("x" & extensionStr), _
                                    "")
                        End If
                    
                    ' clean up the raw phone number without the extension
                    
                        rawPhoneNumberStr = Replace(rawPhoneNumberStr, "001-", "")
                        rawPhoneNumberStr = Replace(rawPhoneNumberStr, "+1", "")
                        rawPhoneNumberStr = Replace(rawPhoneNumberStr, "(", "")
                        rawPhoneNumberStr = Replace(rawPhoneNumberStr, ")", "")
                        rawPhoneNumberStr = Replace(rawPhoneNumberStr, "-", "")
                        rawPhoneNumberStr = Replace(rawPhoneNumberStr, ".", "")
                        
                        If (Len(rawPhoneNumberStr) = 10) Then
                            phoneNumberStr = "(" & _
                                Mid(rawPhoneNumberStr, 1, 3) & ") " & _
                                Mid(rawPhoneNumberStr, 4, 3) & "-" & _
                                Mid(rawPhoneNumberStr, 7, 4)
                        End If
                    
                    ' output the results
                    
                        Cells(iCount, 2).Value = phoneNumberStr
                        
                        Cells(iCount, 3).NumberFormat = "@"
                        Cells(iCount, 3).Value = extensionStr

                    
            End If
        Next iCount
               
        
    MsgBox ("Process Complete")
    
End Sub
