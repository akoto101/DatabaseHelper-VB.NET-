Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Public Class Database
    Private Shared Function Fields(ByVal Data As Dictionary(Of String, Object), ByVal num As Integer) As String
        Dim builder As StringBuilder = New StringBuilder()

        If num = 0 Then

            For Each key In Data
                builder.Append(key.Key & ",")
            Next
        ElseIf num = 2 Then

            For Each key In Data
                builder.Append(key.Key & "=" & "?" & ",")
            Next
        Else

            For Each key In Fields
                builder.Append("?" & ",")
            Next
        End If

        Return builder.ToString().TrimEnd(","c)
    End Function

    Private Shared Sub Parameter(ByVal command As IDbCommand, ByVal param As Dictionary(Of String, Object))
        For Each p In param
            Dim pa = command.CreateParameter()
            pa.ParameterName = p.Key
            pa.Value = p.Value
            command.Parameters.Add(pa)
        Next

        command.ExecuteNonQuery()
    End Sub

    Private GLOBAL_CONNECTION As IDbConnection
    Private command As IDbCommand
    Private adapter As IDbDataAdapter

    Public Sub New(ByVal connection As IDbConnection, ByVal Command As IDbCommand)
        GLOBAL_CONNECTION = connection
        GLOBAL_CONNECTION.ConnectionString = connection.ConnectionString
        Command = Command
    End Sub

    Private Sub Init()
        GLOBAL_CONNECTION.Open()
        command.Connection = GLOBAL_CONNECTION
    End Sub

    Public Sub setAdapter(ByVal Adapter As IDbDataAdapter)
        Adapter = Adapter
    End Sub

    Public Sub Insert(ByVal Table_Name As String, ByVal Data As Dictionary(Of String, Object))
        Try
            Init()
            command.CommandText = $"INSERT INTO {Table_Name} ({Fields(Data, 0)}) VALUES ({Fields(Data, 1)})"
            Parameter(command, Data)
        Catch ex As Exception
            Throw ex
        Finally
            GLOBAL_CONNECTION.Close()
        End Try
    End Sub

    Public Sub Bulk_Insert(ByVal Table_Name As String, ByVal Data As Dictionary(Of String, Object()))
        Try
            Init()
            Dim i As Integer = 0

            For Each Key In Data.Keys

                For Each dataSet In Data.Values
                    command.CommandText = $"INSERT INTO {Table_Name} ({Key}) VALUES (""{dataSet(i)}"")"
                    command.ExecuteNonQuery()
                    i += 1
                Next
            Next

        Catch ex As Exception
            Throw ex
        Finally
            GLOBAL_CONNECTION.Close()
        End Try
    End Sub

    Public Sub Delete(ByVal Table_Name As String, ByVal Data As Dictionary(Of String, Object))
        Try
            Init()
            command.Connection = GLOBAL_CONNECTION
            command.CommandText = $"DELETE FROM {Table_Name} WHERE {Fields(Data, 0)}={Fields(Data, 2)}"
            Parameter(command, Data)
        Catch ex As Exception
            Throw ex
        Finally
            GLOBAL_CONNECTION.Close()
        End Try
    End Sub

    Public Sub Delete(ByVal Table_Name As String, ByVal Condition As String)
        Try
            Init()
            Dim builder As StringBuilder = New StringBuilder()

            If Condition <> "" Then
                builder.Append($"DELETE FROM {Table_Name} ")
                builder.Append($"WHERE {Condition}")
            Else
                builder.Append($"DELETE {Table_Name} ")
            End If

            command.CommandText = builder.ToString()
            command.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            GLOBAL_CONNECTION.Close()
        End Try
    End Sub

    Public Sub Update(ByVal Table_Name As String, ByVal Data As Dictionary(Of String, Object), ByVal Condition As String)
        Try
            Init()
            Dim builder As StringBuilder = New StringBuilder()
            builder.Append($"UPDATE {Table_Name} SET {Fields(Data, 2)} ")

            If Condition <> "" Then
                builder.Append($"WHERE {Condition}")
            End If

            command.CommandText = builder.ToString()
            Parameter(command, Data)
        Catch ex As Exception
            Throw ex
        Finally
            GLOBAL_CONNECTION.Close()
        End Try
    End Sub

    Public Function [Select](ByVal Table_Name As String, ByVal Data As Dictionary(Of String, Object), ByVal Condition As String) As DataTable
        Dim ds As DataSet = New DataSet()

        Try
            Init()
            Dim builder As StringBuilder = New StringBuilder()

            If Data Is Nothing Then
                builder.Append($"SELECT * from {Table_Name} ")

                If Condition <> "" Then
                    builder.Append($"WHERE {Condition}")
                End If

                command.CommandText = builder.ToString()
            Else
                builder.Append($"SELECT {Fields(Data, 0)} from {Table_Name} ")

                If Condition <> "" Then
                    builder.Append($"WHERE {Condition}")
                End If

                command.CommandText = builder.ToString()
            End If

            adapter.SelectCommand = command
            adapter.Fill(ds)
        Catch ex As Exception
            Throw ex
        Finally
            GLOBAL_CONNECTION.Close()
        End Try

        Return ds.Tables(0)
    End Function
End Class
