Sub CopyFilesBasedOnNameUSE()
    Dim xRg As Range, xCell As Range
    Dim xFileDlg As FileDialog
    Dim xSelectedItems As FileDialogSelectedItems
    Dim xVal As String
    Dim notCopied As String
    Dim successfullyCopied As Integer
    Dim destinationDict As Object
    Dim ws As Worksheet
    Dim folderPath As String
    Dim response As VbMsgBoxResult
    Dim fixedPath As String
    Dim accountName As String
    Dim replaceCount As Long
    Dim defaultPath As String
    Dim defaultPath2 As String
    
    ' Set the default folder path for all dialog boxes
    defaultPath = "PATH OF WHERE STATEMENTS ARE HELD"
    defaultPath2 = "PATH OF MAIN FOLDER WHERE STATEMENTS ARE GETTING COPIED TO"
    
    ' Set the worksheet to the one where the button is located
    Set ws = ActiveSheet

    ' Define the range in column A starting from row 2
    Set xRg = ws.Range("A2", ws.Cells(ws.Rows.Count, 1).End(xlUp))
    
    ' Ask for the new text once
    Dim startPos As Integer, substringLength As Integer
    startPos = 11 ' Starting position for the substring
    substringLength = 6 ' Length of the substring to replace
    Dim newText As String
    newText = InputBox("Please enter the date of the statement needing to be copied (MMDDYY):")
    
    If Len(newText) <> substringLength Then
        MsgBox "Please enter exactly " & substringLength & " characters for the replacement text.", vbExclamation
        Exit Sub
    End If

    ' Replace the substring in column A for each cell
    replaceCount = 0
    For Each xCell In xRg
        Dim originalText As String
        originalText = xCell.Value
        If Len(originalText) >= 15 Then
            xCell.Value = Left(originalText, startPos - 1) & newText & Mid(originalText, startPos + substringLength)
            replaceCount = replaceCount + 1
        End If
    Next xCell

    ' Initialize destination dictionary for later file copying
    Set destinationDict = CreateObject("Scripting.Dictionary")
    
    ' Loop through each row to verify or prompt for destination links
    For Each xCell In xRg.Rows
        If xCell.Cells(1, 1).Value <> "" Then
            folderPath = xCell.Cells(1, 2).Value
            Dim accountNumber As String
            accountNumber = xCell.Cells(1, 3).Value
            accountName = xCell.Cells(1, 4).Value
            
            If Len(Dir(folderPath, vbDirectory)) = 0 Then
                response = MsgBox("Broken link found for:" & vbCrLf & _
                                  "Account Number: " & accountNumber & vbCrLf & _
                                  "Account Name: " & accountName & vbCrLf & _
                                  "Do you want to fix it?", vbYesNoCancel + vbExclamation)
                
                If response = vbYes Then
                    Set xFileDlg = Application.FileDialog(msoFileDialogFolderPicker)
                    xFileDlg.InitialFileName = defaultPath2
                    xFileDlg.Title = "Please choose a new destination folder for " & accountNumber
                    If xFileDlg.Show = -1 Then
                        fixedPath = xFileDlg.SelectedItems(1)
                        xCell.Cells(1, 2).Value = fixedPath
                        destinationDict.Add xCell.Cells(1, 1).Value, fixedPath
                    Else
                        MsgBox "No folder selected. Link remains broken for Account Number: " & accountNumber & _
                               vbCrLf & "Account Name: " & accountName
                    End If
                ElseIf response = vbCancel Then
                    Exit Sub
                End If
            Else
                destinationDict.Add xCell.Cells(1, 1).Value, folderPath
            End If
        End If
    Next xCell
    
    ' After verifying all links, prompt user to select files for copying
    Set xFileDlg = Application.FileDialog(msoFileDialogFilePicker)
    xFileDlg.InitialFileName = defaultPath
    xFileDlg.Title = "Please select the files to copy"
    xFileDlg.AllowMultiSelect = True
    If xFileDlg.Show <> -1 Then Exit Sub
    Set xSelectedItems = xFileDlg.SelectedItems
    
    ' Initialize variables for tracking copied and not copied files
    notCopied = ""
    successfullyCopied = 0
    
    ' Loop through each selected file to process
    For i = 1 To xSelectedItems.Count
        xVal = xSelectedItems.Item(i)
        Dim xFileName As String
        xFileName = Dir(xVal)
        
        ' Remove .pdf extension if present
        If Right(xFileName, 4) = ".pdf" Then
            xFileName = Left(xFileName, Len(xFileName) - 4)
        End If

        ' Check if the file name is in the dictionary
        If destinationDict.Exists(xFileName) Then
            Dim xDPathStr As String
            xDPathStr = destinationDict(xFileName) & "\"
            
            ' Check if file already exists in the destination folder
            If Dir(xDPathStr & xFileName & ".pdf", vbNormal) = "" Then
                FileCopy xVal, xDPathStr & xFileName & ".pdf"
                successfullyCopied = successfullyCopied + 1
            Else
                notCopied = notCopied & xFileName & ".pdf already exists in the folder" & vbCrLf
            End If
        Else
            ' Add new files to Excel sheet with placeholders for destination and account name
            Dim newRow As Range
            Set newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(1, 0)
            newRow.Value = xFileName
            
            ' Assign first 9 characters as account number in column C
            newRow.Offset(0, 2).Value = Left(xFileName, 9)
            
            ' Prompt user for destination folder and account name
            Set xFileDlg = Application.FileDialog(msoFileDialogFolderPicker)
            xFileDlg.InitialFileName = defaultPath2
            xFileDlg.Title = "Select the destination folder for " & xFileName
            If xFileDlg.Show = -1 Then
                folderPath = xFileDlg.SelectedItems(1)
                newRow.Offset(0, 1).Value = folderPath
                destinationDict.Add xFileName, folderPath
            End If
            
            ' Add account name
            accountName = InputBox("Enter the account name for " & xFileName & ":")
            newRow.Offset(0, 3).Value = accountName
            
            ' Copy the file to the new destination
            If Len(folderPath) > 0 Then
                If Dir(folderPath & "\" & xFileName & ".pdf", vbNormal) = "" Then
                    FileCopy xVal, folderPath & "\" & xFileName & ".pdf"
                    successfullyCopied = successfullyCopied + 1
                Else
                    notCopied = notCopied & xFileName & ".pdf already exists in the folder" & vbCrLf
                End If
            End If
        End If
    Next i
    
    ' Display message with results
    If notCopied = "" Then
        MsgBox "All files were copied successfully.", vbInformation
    Else
        MsgBox successfullyCopied & " files were copied successfully." & vbCrLf & _
               "The following files were not copied:" & vbCrLf & notCopied, vbExclamation
    End If
End Sub

