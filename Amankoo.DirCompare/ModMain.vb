Module Main

    Sub Main()
        Dim arguments As String() = Environment.GetCommandLineArgs()

        'If processing the parameters returns false, something didn't work: Show help and return error code
        If Parameter.ProcessParameter(arguments) = False Then
            Help.ShowHelp()
            Environment.Exit(ErrorCode.ErrorCode)
        End If
#If DEBUG Then
        For i = 0 To arguments.Count - 1
            Debug.WriteLine("[main] arguments[" & i & "]: " & arguments(i))
        Next
#End If
        Parameter.Started = DateTime.Now
        Output.WriteHeader()
        Compare.ComparePath("")
        Parameter.Ended = DateTime.Now
        Output.WriteFooter()
    End Sub

End Module
