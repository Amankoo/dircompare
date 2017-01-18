Module Output

    Sub WriteHeader()
        If Parameter.NoHeader = True Then Exit Sub
        Console.WriteLine("Amankoo Compare")
        Console.WriteLine("Version " & System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString)
        Console.WriteLine("################################################################################")
        Parameter.WriteParametersToConsole()
        Console.WriteLine("################################################################################")
        Console.WriteLine("Started: " & Parameter.Started)
    End Sub

    Sub WriteFooter()
        If Parameter.NoFooter = True Then Exit Sub
        Console.WriteLine("################################################################################")
        Console.WriteLine("Ended: " & Parameter.Ended)
        Console.WriteLine("Total files: " & Counter.Files)
        Console.WriteLine("Identical files: " & Counter.Files_Identical)
        Console.WriteLine("Different files: " & Counter.Files_Different)
        Console.WriteLine("Missing files: " & Counter.Files_Missing)
        Console.WriteLine("")
        Console.WriteLine("Total run time: " & DateDiff(DateInterval.Second, Parameter.Started, Parameter.Ended) & " seconds")
        Console.ReadLine()
    End Sub
End Module
