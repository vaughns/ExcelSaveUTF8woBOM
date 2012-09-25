Option Explicit

'*********************
' Module: modSaveUTF8
' Description:
'   contains functions to save data from active workbook as UTF-8, without BOM
' Required References:
'   Visual Basic For Applications
'   Microsoft Excel 11.0 Object Library
'   Microsoft ActiveX Data Objects 2.5 Library
'
'*********************

Sub SaveUTF8woBOM()

''''''''''''''''''''''
' Sub: SaveUTF8woBOM
' Process:
'   Opens save file dialog to get a path for saving the file
'   Copies the active sheet to a new workbook
'   Saves the new workbook in Unicode format to the path given
'   Closes the new workbook (thus avoiding permission conflict)
'   Calls UnicodeToUTF8woBOM on the file saved to convert it from Unicode to UTF-8 without BOM
'   Tells user that file is saved
'
''''''''''''''''''''''

    Dim wbCurr As Workbook  'this is the active workbook, from which you are saving the data
    Dim wbTemp As Workbook  'this is a temporary workbook that becomes your saved file
    Dim sFileSavePath As String
    Dim iSave As Integer 'this is for error checking in case the file already exists

    Set wbCurr = ActiveWorkbook

' This goes through the save/file exists/overwrite process in the way users are used to
    iSave = vbNo
    While iSave = vbNo
' Get a save path from the user
        sFileSavePath = Application.GetSaveAsFilename("", _
                        "UTF-8 w/o BOM text (*.dat), *.dat", , _
                        "Save as UTF-8 without BOM")
        ' if the user presses cancel, end processing
        If sFileSavePath = "False" Then
            MsgBox "Export cancelled."
            Exit Sub
        End If
' Check if file exists
        If Dir(sFileSavePath) <> "" Then
            ' if it does, ask if the user wants to save anyway.
            iSave = MsgBox("The file '" & sFileSavePath & "' already exists. Do you want to replace the existing file?", _
                                vbExclamation + vbYesNo + vbDefaultButton2)
        Else
            ' if it doesn't, then it's okay to save
            iSave = vbYes
        End If
    Wend

' Copy the active sheet to a new workbook
    ActiveSheet.Copy
    Set wbTemp = ActiveWorkbook
    wbTemp.Windows(1).Visible = False
    wbCurr.Activate

' Save it in Unicode format, which is UTF-16 (or UCS-2, Little Endian)
' Since we've already asked the user about conflicts, we can block the system's asking the user about them
    Application.DisplayAlerts = False
    wbTemp.SaveAs Filename:=sFileSavePath, FileFormat:=xlUnicodeText
    Application.DisplayAlerts = True
' Close it
    wbTemp.Close SaveChanges:=False

' Clean up
    Set wbCurr = Nothing
    Set wbTemp = Nothing

' Pass it to conversion function, then let the user know we're done.
    sFileSavePath = UnicodeToUTF8woBOM(sFileSavePath)
    MsgBox sFileSavePath & " saved."
End Sub

Function UnicodeToUTF8woBOM(sFilePath) As String

''''''''''''''''''''''
' Function: UnicodeToUTF8woBOM
' Arguments: sFilePath
' Process:
'   Opens Unicode file (at sFilePath) with ADODB stream object (strReader)
'   Copies to intermediary UTF-8 stream object (strMiddle)
'   Skips first three bytes (the BOM), then copies the rest to final UTF-8 stream object (strWriter)
'   Saves file from final stream to sFilePath, replacing input file
'   Returns file path where saved (sFilePath)
''''''''''''''''''''''

' Create strReader, strMiddle, & strWriter streams
    Dim strReader As Object, strMiddle As Object, strWriter As Object
    Set strReader = CreateObject("Adodb.Stream")
    Set strMiddle = CreateObject("Adodb.Stream")
    Set strWriter = CreateObject("Adodb.Stream")

' Load from the text file
    strReader.Type = adTypeText
    strReader.Mode = adModeReadWrite
    ' This is where you set the charset. You have to declare it, as it defaults to UTF-8
    ' The computer will handle all charsets in the registry here:
    ' HKEY_CLASSES_ROOT\MIME\Database\Charset
    strReader.Charset = "Unicode"
    strReader.LineSeparator = adLF
    strReader.Open
    strReader.LoadFromFile sFilePath

' Copy all data from strReader to strMiddle, which converts it to UTF-8
    strMiddle.Mode = adModeReadWrite
    strMiddle.Type = adTypeText
    strMiddle.Charset = "UTF-8"
    strMiddle.Open
    strReader.CopyTo strMiddle
    strReader.Flush
    strMiddle.Flush

' Clean up strReader
    strReader.Close
    Set strReader = Nothing

' Copy data from strMiddle to strWriter, skipping BOM
    strWriter.Mode = adModeReadWrite
    strWriter.Type = adTypeBinary
    strWriter.Open
    'skip the BOM
    strMiddle.Position = 3
    'and then copy the rest
    strMiddle.CopyTo strWriter
    strMiddle.Flush
    strWriter.Flush

' Clean up strMiddle
    strMiddle.Close
    Set strMiddle = Nothing

' Overwrite input file
    strWriter.SaveToFile sFilePath, adSaveCreateOverWrite

' Return file name
    UnicodeToUTF8woBOM = sFilePath

' Clean up strWriter
    strWriter.Close
    Set strWriter = Nothing
End Function


