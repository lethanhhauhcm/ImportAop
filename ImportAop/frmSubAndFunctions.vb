

Module frmSubAndFunctions
    Public pstrConnectionAop As String = "Data Source=118.69.68.197; User Id= user_aop; Password= VietHealthy@170172#; Connection Timeout=30;"
    'Public pstrConnectionAop As String = "Data Source=42.117.5.70; User Id= user_aop; Password= VietHealthy@170172#; Connection Timeout=30;"
    Public pstrConnectionRAS As String = "server=118.69.81.103;uid=user_ras;pwd=VietHealthy@170172#;database=RAS12"
    'Public pstrConnectionRAS As String = "server=118.69.81.103;uid=user_ras;pwd=VietHealthy@170172#;database=RAS12HAN"
    Public pstrPrg As String = "ImportAOP"
    Public pobjSqlAop As New clsTvcs
    Public pobjSqlRas As New clsTvcs
    Public pobjOdbc As New clsOdbc
    Public pstrCity As String = "SGN"
    'Public pstrCity As String = "HAN"
    Public Function Append2TextFile(ByVal strText As String) As Boolean
        Dim strLogFile As String = My.Application.Info.DirectoryPath & "\" _
                                            & Format(Today, "yyMMdd") & pstrPrg & ".txt"

        Dim objLogFile As New System.IO.StreamWriter(strLogFile, True)
        objLogFile.WriteLine(strText)
        objLogFile.Close()
        objLogFile = Nothing
        Return True
    End Function
End Module
