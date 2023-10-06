Imports MySql.Data.MySqlClient

Public Class PredictiveDBInput

    Dim MysqlConn As MySqlConnection
    Dim command1 As MySqlCommand
    Dim command2 As MySqlCommand
    Dim command3 As MySqlCommand

#Region "AHU"

    Sub AHUPartCount(ByVal partname As String)

        MysqlConn = New MySqlConnection With {.ConnectionString = "server = 127.0.0.1; userid = root; password = ; database = airpacpredictivedb; SslMode = None"}
        Dim reader As MySqlDataReader
        Dim query1 As String
        Dim query2 As String
        Dim query3 As String

        MysqlConn.Open()
        query1 = "SELECT * FROM ahu_partcount WHERE PartName = '" & partname & "'"
        command1 = New MySqlCommand(query1, MysqlConn)
        reader = command1.ExecuteReader
        If reader.HasRows Then
            MysqlConn.Close()
            MysqlConn.Open()
            query2 = "update airpacpredictivedb.ahu_partcount set Count = Count + 1 where PartName = '" & partname & "'"
            command2 = New MySqlCommand(query2, MysqlConn)
            reader = command2.ExecuteReader
        Else
            MysqlConn.Close()
            MysqlConn.Open()
            query3 = "insert into airpacpredictivedb.ahu_partcount (PartName, Count) values ('" & partname & "', 1)"
            command3 = New MySqlCommand(query3, MysqlConn)
            reader = command3.ExecuteReader
        End If

        MysqlConn.Close()

        MysqlConn.Dispose()

    End Sub

    Sub AHUFanCount(ByVal partname As String)

        MysqlConn = New MySqlConnection With {.ConnectionString = "server = 127.0.0.1; userid = root; password = ; database = airpacpredictivedb; SslMode = None"}
        Dim reader As MySqlDataReader
        Dim query1 As String
        Dim query2 As String
        Dim query3 As String

        MysqlConn.Open()
        query1 = "select * from airpacpredictivedb.ahu_fancount where FanArticleNo = '" & partname & "'"
        command1 = New MySqlCommand(query1, MysqlConn)
        reader = command1.ExecuteReader
        If reader.HasRows Then
            MysqlConn.Close()
            MysqlConn.Open()
            query2 = "update airpacpredictivedb.ahu_fancount set Count = Count + 1 where FanArticleNo = '" & partname & "'"
            command2 = New MySqlCommand(query2, MysqlConn)
            reader = command2.ExecuteReader
        Else
            MysqlConn.Close()
            MysqlConn.Open()
            query3 = "insert into airpacpredictivedb.ahu_fancount (FanArticleNo, Count) values ('" & partname & "', 1)"
            command3 = New MySqlCommand(query3, MysqlConn)
            reader = command3.ExecuteReader
        End If

        MysqlConn.Close()

        MysqlConn.Dispose()

    End Sub

#End Region

#Region "Passbox"

    Sub PassboxPartCount(ByVal partname As String)

        MysqlConn = New MySqlConnection With {.ConnectionString = "server = 127.0.0.1; userid = root; password = ; database = airpacpredictivedb; SslMode = None"}
        Dim reader As MySqlDataReader
        Dim query1 As String
        Dim query2 As String
        Dim query3 As String

        MysqlConn.Open()
        query1 = "select * from airpacpredictivedb.passbox_partcount where PartName = '" & partname & "'"
        command1 = New MySqlCommand(query1, MysqlConn)
        reader = command1.ExecuteReader
        If reader.HasRows Then
            MysqlConn.Close()
            MysqlConn.Open()
            query2 = "update airpacpredictivedb.passbox_partcount set Count = Count + 1 where PartName = '" & partname & "'"
            command2 = New MySqlCommand(query2, MysqlConn)
            reader = command2.ExecuteReader
        Else
            MysqlConn.Close()
            MysqlConn.Open()
            query3 = "insert into airpacpredictivedb.passbox_partcount (PartName, Count) values ('" & partname & "', 1)"
            command3 = New MySqlCommand(query3, MysqlConn)
            reader = command3.ExecuteReader
        End If

        MysqlConn.Close()

        MysqlConn.Dispose()

    End Sub

    Sub PassboxProjectCount(ByVal pbProjNo As String, ByVal pbSize As String, ByVal pbType As String, ByVal pbConstruction As String, ByVal Door2Side As String)

        MysqlConn = New MySqlConnection With {.ConnectionString = "server = 127.0.0.1; userid = root; password = ; database = airpacpredictivedb; SslMode = None"}
        Dim reader As MySqlDataReader
        Dim query As String

        Dim DoorConfig As String = "Straight"
        If Door2Side <> "BACK" Then
            DoorConfig = "L-Type"
        End If

        MysqlConn.Open()

        query = "SELECT COUNT(*) FROM `passbox_projectlist` WHERE ProjectNo = " & pbProjNo
        command3 = New MySqlCommand(query, MysqlConn)
        Dim rowCount As Integer = command3.ExecuteScalar

        If rowCount = 0 Then
            query = "INSERT INTO passbox_projectlist (ProjectNo, InternalSize, OperationType, ConstructionType, DoorConfiguration)
                     VALUES ('" & pbProjNo & "', '" & pbSize & "', '" & pbType & "', '" & pbConstruction & "', '" & DoorConfig & "')"
            command3 = New MySqlCommand(query, MysqlConn)
            reader = command3.ExecuteReader
            reader.Close()
        Else
            query = "UPDATE passbox_projectlist 
                     SET InternalSize = '" & pbSize & "', 
                         OperationType = '" & pbType & "', 
                         OperationType = '" & pbConstruction & "',
                         DoorConfiguration = '" & DoorConfig & "'
                     WHERE ProjectNo = " & pbProjNo & ""
            command3 = New MySqlCommand(query, MysqlConn)
            reader = command3.ExecuteReader
        End If

        MysqlConn.Close()

        MysqlConn.Dispose()

    End Sub

#End Region

End Class