Module ErrorCode
    Public Const ERR_NO_ERROR As Byte = 0
    Public Const ERR_SOURCEPATH_NOT_FOUND As Byte = 1
    Public Const ERR_TARGETPATH_NOT_FOUND As Byte = 2
    Public Const ERR_FILESPARAMETER_NOT_FOUND As Byte = 3
    Public Const ERR_FCPARAMETER_NOT_FOUND As Byte = 4
    Public Const ERR_FILELISTINGPARAMETER_NOT_FOUND As Byte = 5
    Private BytErrorCode As Byte = 0
    Public Property ErrorCode() As Byte
        Get
            Return BytErrorCode
        End Get
        Set(ByVal value As Byte)
            BytErrorCode = value
        End Set
    End Property
End Module
