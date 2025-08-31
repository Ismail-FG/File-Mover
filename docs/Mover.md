## Mover
```groovy
Function PindahkanDokumen()

    Dim berkasAsal As String
    Dim berkasTujuan As String
    Dim daftarDok As Variant
    Dim namalokDok As String
    Dim i As Integer
    Dim ekstensiDok As String
    
    berkasAsal = Range("B1").Value
    berkasTujuan = Range("B2").Value
    ekstensiDok = Range("B3").Value
    
    daftarDok = GetFileList(berkasAsal)
    
    ' Cek apakah daftarDok kosong
    If IsEmpty(daftarDok) Or (UBound(daftarDok) = 0 And daftarDok(0) = "") Then
        Debug.Print "Tidak ada file di folder asal."
        Exit Function
    End If
    
    Debug.Print "Berkas Asal: " & berkasAsal
    Debug.Print "Berkas Tujuan: " & berkasTujuan
    
    For i = 0 To UBound(daftarDok)
    
        namalokDok = berkasAsal & "\" & daftarDok(i)
        Debug.Print namalokDok
        
        If Right(namalokDok, Len(ekstensiDok) + 1) = "." & ekstensiDok Then

            MoveOldFiles namalokDok, berkasTujuan

        End If
    
    Next i
    

End Function

Function MoveOldFiles(ByVal filePath As String, ByVal folderDestination As String)
    Dim sourceFile As String
    Dim destinationFile As String
    Dim fileSystem As Object
    Dim fileObj As Object
    Dim fileAge As Double
    
    ' Set file paths
    sourceFile = filePath
    destinationFile = folderDestination & "\" & Dir(filePath) ' Ensure full file path
    
    ' Create FileSystemObject
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    
    ' Check if file exists
    If fileSystem.FileExists(sourceFile) Then
        ' Get file object
        Set fileObj = fileSystem.GetFile(sourceFile)
        
        ' Calculate file age in minutes
        fileAge = ((Now - fileObj.DateLastModified) * 1440)
        
        ' Move file if older than 60 minutes
        If fileAge > 90 Then
            fileObj.Move destinationFile
            Debug.Print "File moved successfully from " & sourceFile & " to " & destinationFile
        Else
            Debug.Print "File is not old enough to move."
        End If
    Else
        Debug.Print "File not found!"
    End If
    
    ' Cleanup
    Set fileObj = Nothing
    Set fileSystem = Nothing
End Function

Function GetFileList(folderPath As String) As Variant
    Dim fileSystem As Object
    Dim file As Object
    Dim fileList() As String
    Dim i As Integer
    
    ' Create FileSystemObject
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the folder exists
    If Not fileSystem.FolderExists(folderPath) Then
        MsgBox "Folder not found!", vbExclamation
        Exit Function
    End If
    
    ' Initialize counter and array
    i = 0
    ReDim fileList(0)

    ' Loop through files in the folder
    For Each file In fileSystem.GetFolder(folderPath).Files
        ' Resize array dynamically
        ReDim Preserve fileList(i)
        fileList(i) = file.Name
        i = i + 1
    Next file
    
    ' Return array
    GetFileList = fileList
    
    ' Cleanup
    Set fileSystem = Nothing
End Function

