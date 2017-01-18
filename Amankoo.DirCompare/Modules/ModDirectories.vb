Imports System.IO

Module Directories
    ''' <summary>
    ''' Checks if the local directory exists
    ''' </summary>
    ''' <param name="DirectoryName">Path of a local directory (No UNC path)</param>
    ''' <returns><see langword="true"/> if the path exists. Otherwise <see langword="false"/></returns>
    Function LocalDirectoryExists(DirectoryName) As Boolean
        If Directory.Exists(DirectoryName) Then
            Return True
        Else
            Return False
        End If
    End Function

    Function GetSubDirectories(Path As String) As DirectoryInfo()
        Dim DirInfo As New DirectoryInfo(Path)
        Dim Dirs() As DirectoryInfo = DirInfo.GetDirectories()
        Return Dirs
    End Function

    Function GetSubFiles(Path As String) As FileInfo()
        Dim DirInfo As New DirectoryInfo(Path)
        Dim Files() As FileInfo = DirInfo.GetFiles(Parameter.SearchPattern)
        Return Files
    End Function
End Module
