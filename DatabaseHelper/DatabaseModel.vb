Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Public Class DatabaseModel
    Inherits Dictionary(Of String, Object)

    Private field As String
    Private value As Object

    Public Sub New()
    End Sub

    Public Sub Bulk_Add(ByVal Field As String(), ByVal Value As Object())
        For i As Integer = 0 To Field.Length - 1
            Me.Add(Field(i), Value(i))
        Next
    End Sub

    Public Property Fields As String
        Get
            Return field
        End Get
        Set(ByVal value As String)
            field = value
        End Set
    End Property

    Public Property Values As Object
        Get
            Return value
        End Get
        Set(ByVal value As Object)
            Me.value = value
        End Set
    End Property
End Class
