Imports System.Data.Odbc
Module Module1
    Public con As New OdbcConnection
    Public da As OdbcDataAdapter
    Public ds As DataSet
    Public rd As OdbcDataReader
    Public cmd As OdbcCommand
    Public myDb As String

    Public Sub koneksi()
        myDb = "Driver={Mysql ODBC 3.51 Driver};database=db_penjualan;server=localhost;uid=root"
        con = New OdbcConnection(myDb)
        If con.State = ConnectionState.Closed Then con.Open()
    End Sub

End Module
