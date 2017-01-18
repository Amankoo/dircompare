Module Functions
    Function HashToString(Hash() As Byte)
        Dim HashValue As String = ""

        For i = 0 To Hash.Length - 1
            HashValue += Hash(i).ToString("X2")
        Next i

        Return HashValue.ToUpper
    End Function
End Module
