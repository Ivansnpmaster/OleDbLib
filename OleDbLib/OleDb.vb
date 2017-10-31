Imports System.Data.OleDb

Public Class OleDb

    'Deve-se prover uma string válida de conexão no construtor
    Dim cS As String = ""

    Sub New(connectionString As String)
        cS = connectionString
    End Sub

    Public Function InsertData(tableName As String, bankColumns() As String, itemsToInsert() As Object) As Boolean

        'String de comando
        Dim sql As String = String.Format("INSERT INTO {0} (", tableName)
        'String para o adicionar de parâmetros no OleDbCommand
        Dim a(bankColumns.Count - 1) As String

        'Montando a string de comando sem o último elemento (i até -2)
        For i As Integer = 0 To bankColumns.Count - 2
            a(i) = "@" + bankColumns(i)
            sql += a(i) + ", "
        Next
        'Adicionando o último elemento
        sql += a(a.Count - 1) + ") VALUES ("

        'Adicionando os parâmetros com o @
        For i As Integer = 0 To a.Count - 2
            sql += a(i) + ", "
        Next
        'Adicionando o último elemento
        sql += a(a.Count - 1) + ")"

        Dim cmd As OleDbCommand = New OleDbCommand(sql)

        'Adicionando os parâmetros aos OleDbCommand
        For i As Integer = 0 To a.Count - 1
            cmd.Parameters.AddWithValue(a(i), itemsToInsert(i))
        Next

        Return ExecuteCommand(cmd)

    End Function

    Public Function DeleteData(tableName As String, bankColumns() As String, operators() As String, itemsToInsert() As Object) As Boolean

        'String de comando
        Dim sql As String = String.Format("DELETE FROM {0} WHERE ", tableName)
        'String para o adicionar de parâmetros no OleDbCommand
        Dim a(bankColumns.Count - 1) As String

        'Montando a string de comando sem o último elemento (i até -2)
        For i As Integer = 0 To bankColumns.Count - 2
            'Adicionando o parâmetro i para poder colocar a mesma coluna da database mais de uma vez no WHERE
            a(i) = String.Format("@{0}{1}", bankColumns(i), i)
            sql += String.Format("{0} {1} {2} AND ", bankColumns(i), operators(i), a(i))
        Next
        'Adicionando o último elemento
        sql += String.Format("{0} {1} {2}", bankColumns(a.Count - 1), operators(a.Count - 1), a(a.Count - 1))

        Dim cmd As OleDbCommand = New OleDbCommand(sql)

        'Adicionando os parâmetros aos OleDbCommand
        For i As Integer = 0 To a.Count - 1
            cmd.Parameters.AddWithValue(a(i), itemsToInsert(i))
        Next

        Return ExecuteCommand(cmd)

    End Function

    Public Function UpdateData(tableName As String, bankColumns() As String, itemsToInsert() As Object, whereBankColumns() As String, operators() As String, whereItems() As Object) As Boolean

        'String de comando
        Dim sql As String = String.Format("UPDATE {0} SET ", tableName)
        'String para o adicionar de parâmetros no OleDbCommand, tamanho de itens com o @
        Dim a(bankColumns.Count - 1 + whereItems.Count - 1) As String

        'Montando a string de comando com os parâmetros novos (bankColumn = @novoItem) sem o último elemento
        For i As Integer = 0 To bankColumns.Count - 2
            a(i) = String.Format("@{0}{1}", bankColumns(i), i)
            sql += String.Format("{0} = {1}, ", bankColumns(i), a(i))
        Next

        Dim j As Integer = bankColumns.Count - 1
        'Adicionando o último elemento
        a(j) = String.Format("@{0}{1}", bankColumns(j), j)
        sql += String.Format("{0} = {1} WHERE ", bankColumns(j), a(j))

        'Montando a string de comando com os parâmetros WHERE (whereColumn operador @parâmetro) sem o último elemento
        For i As Integer = bankColumns.Count To a.Count - 2
            a(i) = String.Format("@{0}{1}", whereBankColumns(i - bankColumns.Count), i)
            sql += String.Format("{0} {1} {2}, ", whereBankColumns(i - bankColumns.Count), operators(i - bankColumns.Count), a(i))
        Next

        j = a.Count - 1 - bankColumns.Count
        'Adicionando o último elemento
        a(j) = String.Format("@{0}{1}", whereBankColumns(j), a.Count - 1)
        sql += String.Format("{0} {1} {2}, ", whereBankColumns(j), operators(j), a(a.Count - 1))

        Dim cmd As OleDbCommand = New OleDbCommand(sql)

        'Adicionando os parâmetros aos OleDbCommand
        For i As Integer = 0 To bankColumns.Count - 1
            cmd.Parameters.AddWithValue(a(i), bankColumns(i))
        Next
        For i As Integer = bankColumns.Count To a.Count - 1
            cmd.Parameters.AddWithValue(a(i), whereItems(i - bankColumns.Count))
        Next

        Return ExecuteCommand(cmd)

    End Function

    Public Function SelectData(tableName As String, bankColumns() As String, whereBankColumns() As String, operators() As String, whereItems() As Object) As DataTable

        'String de comando
        Dim sql As String = "SELECT "
        'String para o adicionar de parâmetros no OleDbCommand
        Dim a(whereBankColumns.Count - 1) As String

        'Montando a string de comando sem o último elemento (i até -2)
        For i As Integer = 0 To bankColumns.Count - 2
            sql += bankColumns(i) + ", "
        Next
        'Adicionando o último elemento
        sql += String.Format("{0} FROM {1} WHERE", bankColumns(bankColumns.Count - 1), tableName)

        'Adicionando na string de comando as condições com @'s
        For i As Integer = 0 To whereBankColumns.Count - 2
            a(i) = String.Format("@{0}{1}", whereBankColumns, i)
            sql += String.Format("{0} {1} {2}", whereBankColumns(i), operators(i), a(i))
        Next

        Dim j As Integer = whereBankColumns.Count - 1
        'Adicionando o último elemento
        a(j) = String.Format("@{0}{1}", whereBankColumns, j)
        sql += String.Format("{0} {1} {2}", whereBankColumns(j), operators(j), a(j))

        Dim dt As DataTable = New DataTable

        Dim con As OleDbConnection = New OleDbConnection(cS)
        Dim cmd As OleDbCommand = New OleDbCommand(sql, con)

        'Adicionando os parâmetros aos OleDbCommand
        For i As Integer = 0 To whereBankColumns.Count - 1
            cmd.Parameters.AddWithValue(a(i), whereItems(i))
        Next

        Try
            Dim da As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            'Fazendo o OleDbAdapter preencher o DataTable
            da.Fill(dt)

        Catch ex As Exception
        Finally
            con.Close()
        End Try
        
        Return dt

    End Function

    Private Function ExecuteCommand(cmd As OleDbCommand) As Boolean

        Dim success As Boolean = True
        Dim con As OleDbConnection = New OleDbConnection(cS)

        cmd.CommandType = CommandType.Text
        cmd.Connection = con

        Try
            'Abrindo a conexão
            con.Open()
            'Executando o comando sem retorno
            cmd.ExecuteNonQuery()

        Catch ex As Exception
            success = False
        Finally
            con.Dispose()
        End Try

        Return success

    End Function

End Class