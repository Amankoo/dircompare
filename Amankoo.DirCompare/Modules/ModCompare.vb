Imports System.IO
Imports System.Security.Cryptography

Module Compare
    Sub ComparePath(Path As String)
        Dim SubFiles As FileInfo()
        Dim SubFolders As DirectoryInfo()
        Dim CurrentSourcePath As String
        Dim CurrentTargetPath As String
        Dim TargetFilePath As String
        Dim CurrentPath As String

        If Path <> "" Then
            CurrentSourcePath = Parameter.PathSource & "\" & Path
            CurrentTargetPath = Parameter.PathTarget & "\" & Path
            CurrentPath = Path & "\"
        Else
            CurrentSourcePath = Parameter.PathSource
            CurrentTargetPath = Parameter.PathTarget
            CurrentPath = ""
        End If


        SubFolders = Directories.GetSubDirectories(CurrentSourcePath)
        SubFiles = Directories.GetSubFiles(CurrentSourcePath)
        Console.WriteLine("")
        Console.WriteLine("--------------------------------------------------------------------------------")
        Console.WriteLine("Checking directory: " & CurrentSourcePath)
        If SubFiles.Length > 0 Then
            For Each CurrentFile As FileInfo In SubFiles
                Counter.Files += 1
                Dim FileStatus As String
                FileStatus = "File:" & vbTab
                TargetFilePath = CurrentTargetPath & "\" & CurrentFile.Name
                If File.Exists(TargetFilePath) = True Then
                    If CompareFile(Path, CurrentFile.Name) = True Then
                        FileStatus = FileStatus & "Same     "
                        Counter.Files_Identical += 1
                    Else
                        FileStatus = FileStatus & "Different"
                        Counter.Files_Different += 1
                    End If
                Else
                    FileStatus = FileStatus & "Missing  "
                    Counter.Files_Missing += 1
                End If
                FileStatus = FileStatus & vbTab & CurrentFile.Name
                Console.WriteLine(FileStatus)
            Next
        End If

        If SubFolders.Length > 0 Then
            For Each CurrentFolder As DirectoryInfo In SubFolders
                Console.WriteLine("Folder:" & CurrentFolder.Name)
                ComparePath(CurrentPath & CurrentFolder.Name)
            Next
        End If

    End Sub

    Function CompareFile(Path As String, Filename As String) As Boolean
        If Parameter.FileComparisonParameter.Count = 0 Then Return ""

        Dim SourceFileName As String
        Dim TargetFileName As String
        Dim FilesAreIdentical As Boolean = True



        If Path <> "" Then
            SourceFileName = Parameter.PathSource & "\" & Path & "\" & Filename
            TargetFileName = Parameter.PathTarget & "\" & Path & "\" & Filename
        Else
            SourceFileName = Parameter.PathSource & "\" & Filename
            TargetFileName = Parameter.PathTarget & "\" & Filename
        End If

        For i = 0 To Parameter.FileComparisonParameter.Count - 1
            Select Case Parameter.FileComparisonParameter(i)
                Case = FileComparisonMethod.Size
                    Dim SourceFileInfo As New FileInfo(SourceFileName)
                    Dim TargetFileInfo As New FileInfo(TargetFileName)
                    If SourceFileInfo.Length = TargetFileInfo.Length Then
                        FilesAreIdentical = True
                    Else
                        Return False
                    End If
                Case = FileComparisonMethod.MD5
                    Dim SourceFileStream As FileStream = File.OpenRead(SourceFileName)
                    Dim TargetFileStream As FileStream = File.OpenRead(TargetFileName)
                    Dim SourceHash As MD5
                    Dim TargetHash As MD5
                    SourceHash = MD5.Create
                    TargetHash = MD5.Create
                    SourceHash.ComputeHash(SourceFileStream)
                    TargetHash.ComputeHash(TargetFileStream)
                    SourceFileStream.Close()
                    TargetFileStream.Close()
                    If Functions.HashToString(SourceHash.Hash) = Functions.HashToString(TargetHash.Hash) Then
                        FilesAreIdentical = True
                    Else
                        Return False
                    End If
                Case = FileComparisonMethod.SHA1
                    Dim SourceFileStream As FileStream = File.OpenRead(SourceFileName)
                    Dim TargetFileStream As FileStream = File.OpenRead(TargetFileName)
                    Dim SourceHash As SHA1
                    Dim TargetHash As SHA1
                    SourceHash = SHA1.Create()
                    TargetHash = SHA1.Create()
                    SourceHash.ComputeHash(SourceFileStream)
                    TargetHash.ComputeHash(TargetFileStream)
                    SourceFileStream.Close()
                    TargetFileStream.Close()
                    If Functions.HashToString(SourceHash.Hash) = Functions.HashToString(TargetHash.Hash) Then
                        FilesAreIdentical = True
                    Else
                        Return False
                    End If
                Case = FileComparisonMethod.SHA256
                    Dim SourceFileStream As FileStream = File.OpenRead(SourceFileName)
                    Dim TargetFileStream As FileStream = File.OpenRead(TargetFileName)
                    Dim SourceHash As SHA256
                    Dim TargetHash As SHA256

                    SourceHash = SHA256.Create
                    TargetHash = SHA256.Create
                    SourceHash.ComputeHash(SourceFileStream)
                    TargetHash.ComputeHash(TargetFileStream)
                    SourceFileStream.Close()
                    TargetFileStream.Close()
                    If Functions.HashToString(SourceHash.Hash) = Functions.HashToString(TargetHash.Hash) Then
                        FilesAreIdentical = True
                    Else
                        Return False
                    End If
                Case = FileComparisonMethod.Content
                    Dim SourceFileByte As Byte
                    Dim TargetFileByte As Byte
                    Dim SourceFileStream As FileStream
                    Dim TargetFileStream As FileStream
                    Dim ContentIdentical As Boolean

                    SourceFileStream = New FileStream(SourceFileName, FileMode.Open)
                    TargetFileStream = New FileStream(TargetFileName, FileMode.Open)

                    Do
                        ' Read one byte from each file.
                        SourceFileByte = SourceFileStream.ReadByte()
                        TargetFileByte = TargetFileStream.ReadByte()
                        If (SourceFileByte = TargetFileByte) Then
                            ContentIdentical = True
                        Else
                            ContentIdentical = False
                            Exit Do
                        End If
                    Loop While (SourceFileByte <> -1)
                    SourceFileStream.Close()
                    TargetFileStream.Close()
                    If ContentIdentical = True Then
                        FilesAreIdentical = True
                    Else
                        Return False
                    End If
            End Select
        Next
        Return True
    End Function
End Module
