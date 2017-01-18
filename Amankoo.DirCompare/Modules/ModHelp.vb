﻿Module Help
    Public Sub ShowHelp()
        Console.WriteLine("Amankoo Directory Compare")
        Console.WriteLine("#########################")
        Console.WriteLine("")
        Console.WriteLine("Compares 2 directories")
        Console.WriteLine("")
        Console.WriteLine("Syntax: DirCompare source target files [options]")
        Console.WriteLine("")
        Console.WriteLine("Parameters:")
        Console.WriteLine("")
        Console.WriteLine("source:                      Source directory")
        Console.WriteLine("target:                      Target directory")
        Console.WriteLine("files:                       Files to select (e.g. *.*)")
        Console.WriteLine("")
        Console.WriteLine("Options:")
        Console.WriteLine("")
        Console.WriteLine("/S                           Compare subdirectories")
        Console.WriteLine("")
        Console.WriteLine("/FC:file comparison flags    Compare files using the following methods:")
        Console.WriteLine("                             size    Size of the files")
        Console.WriteLine("                             md5     MD5 hash of the files")
        Console.WriteLine("                             sha1    SHA-1 hash of the files")
        Console.WriteLine("                             sha256  SHA-256 hash of the files")
        Console.WriteLine("                             content Content of the files")
        Console.WriteLine("                             Seperate methods with ','.")
        Console.WriteLine("                             The comparison will be executed in the order of the flags")
        Console.WriteLine("")
        Console.WriteLine("/NH                          Displays No Header")
        Console.WriteLine("")
        Console.WriteLine("/NF                          Displays No Footer")
    End Sub
End Module
