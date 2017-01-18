Module Parameter
    Private StrSourcePath As String
    Private StrTargetPath As String
    Private StrSearchPattern As String = "*.*"
    Private DateStarted As DateTime
    Private DateEnded As DateTime
    Private BolSubdirectories As Boolean = False
    Private BolNoHeader As Boolean = False
    Private BolNoFooter As Boolean = False
    Private ObjFCParameter As New List(Of FileComparisonMethod)

    Public Property PathSource() As String
        Get
            Return StrSourcePath
        End Get
        Set(ByVal value As String)
            StrSourcePath = value
        End Set
    End Property

    Public Property PathTarget() As String
        Get
            Return StrTargetPath
        End Get
        Set(ByVal value As String)
            StrTargetPath = value
        End Set
    End Property

    Public Property SearchPattern() As String
        Get
            Return StrSearchPattern
        End Get
        Set(ByVal value As String)
            StrSearchPattern = value
        End Set
    End Property

    Public Property Started() As DateTime
        Get
            Return DateStarted
        End Get
        Set(ByVal value As DateTime)
            DateStarted = value
        End Set
    End Property

    Public Property Ended() As DateTime
        Get
            Return DateEnded
        End Get
        Set(ByVal value As DateTime)
            DateEnded = value
        End Set
    End Property

    Public Property Subdirectories() As Boolean
        Get
            Return BolSubdirectories
        End Get
        Set(ByVal value As Boolean)
            BolSubdirectories = value
        End Set
    End Property

    Public Property NoHeader() As Boolean
        Get
            Return BolNoHeader
        End Get
        Set(ByVal value As Boolean)
            BolNoHeader = value
        End Set
    End Property

    Public Property NoFooter() As Boolean
        Get
            Return BolNoFooter
        End Get
        Set(ByVal value As Boolean)
            BolNoFooter = value
        End Set
    End Property

    Public Property FileComparisonParameter() As List(Of FileComparisonMethod)
        Get
            Return ObjFCParameter
        End Get
        Set(ByVal value As List(Of FileComparisonMethod))
            ObjFCParameter = value
        End Set
    End Property

    Public Sub WriteParametersToConsole()
        Console.WriteLine("Parameters:")
        Console.WriteLine("Source path: " & Parameter.PathSource)
        Console.WriteLine("Target path: " & Parameter.PathTarget)
        Console.WriteLine("File search pattern: " & Parameter.SearchPattern)
        Console.WriteLine("Subdirectories: " & Parameter.Subdirectories)
        If Parameter.FileComparisonParameter.Count > 0 Then
            Console.WriteLine("File comparison methods:")
            For i = 0 To Parameter.FileComparisonParameter.Count - 1
                Console.WriteLine(vbTab & " " & FileComparisonParameter(i).ToString)
            Next
        End If
        Console.WriteLine("No Header: " & Parameter.NoHeader.ToString)
        Console.WriteLine("No Footer: " & Parameter.NoFooter.ToString)
    End Sub
    Public Function ProcessParameter(Arguments As String()) As Boolean
        If Arguments.Length < 4 Then Return False 'We need at least 4 Parameter (Exe file, source path, target path, files)

        If Directories.LocalDirectoryExists(Arguments(1)) Then
            PathSource = Arguments(1)
        Else
            ErrorCode.ErrorCode = ErrorCode.ERR_SOURCEPATH_NOT_FOUND
            Return False
        End If

        If Directories.LocalDirectoryExists(Arguments(2)) Then
            ErrorCode.ErrorCode = ErrorCode.ERR_TARGETPATH_NOT_FOUND
            PathTarget = Arguments(2)
        Else
            Return False
        End If

        'We can skip checking the files parameter, as the number of arguments checks if this parameters are present

        SearchPattern = Arguments(3)

        If Arguments.Length > 4 Then
            For i = 4 To Arguments.Length - 1
                If (Strings.UCase(Strings.Left(Arguments(i), 4))) = "/FC:" Then
                    Dim CommandlineFCParam As String()
                    CommandlineFCParam = Strings.Split(Strings.Right(Arguments(i), Arguments(i).Length - 4), ",")
                    If CommandlineFCParam.Length = 0 Then
                        ErrorCode.ErrorCode = ErrorCode.ERR_FCPARAMETER_NOT_FOUND
                        Return False
                    Else
                        Debug.WriteLine("[parameter.processparemeter] File comparison parameter count: " & CommandlineFCParam.Length)
                        Dim FCParameter As New List(Of FileComparisonMethod)
                        For j = 0 To CommandlineFCParam.Length - 1
                            Select Case Strings.UCase(CommandlineFCParam(j))
                                Case = "SIZE"
                                    FCParameter.Add(FileComparisonMethod.Size)
                                Case = "MD5"
                                    FCParameter.Add(FileComparisonMethod.MD5)
                                Case = "SHA1"
                                    FCParameter.Add(FileComparisonMethod.SHA1)
                                Case = "SHA256"
                                    FCParameter.Add(FileComparisonMethod.SHA256)
                                Case = "CONTENT"
                                    FCParameter.Add(FileComparisonMethod.Content)
                            End Select
                        Next
                        Parameter.FileComparisonParameter = FCParameter
                    End If
                ElseIf (Strings.UCase(Arguments(i) = "/S")) Then
                    Subdirectories = True
                ElseIf (Strings.UCase(Arguments(i) = "/NH")) Then
                    Subdirectories = True
                ElseIf (Strings.UCase(Arguments(i) = "/NF")) Then
                    Subdirectories = True
                End If

            Next
        End If

        Return True
    End Function
End Module
