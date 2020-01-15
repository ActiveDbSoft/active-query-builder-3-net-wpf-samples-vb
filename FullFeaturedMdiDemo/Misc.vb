'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Collections
Imports System.Collections.Generic
Imports System.Xml.Serialization
Imports ActiveQueryBuilder.Core
Imports XmlSerializer = ActiveQueryBuilder.Core.Serialization.XmlSerializer

Public Enum ConnectionTypes
        MSSQL
        MSSQLAzure
        MSAccess
        Oracle
        MySQL
        PostgreSQL
        OLEDB
        ODBC
        SQLite
        DB2
        Firebird
        Excel
    End Enum

    <Serializable>
    <XmlInclude(GetType(ConnectionInfo))>
    Public Class ConnectionList
        <XmlElement(Type:=GetType(ConnectionInfo))>
        Private _connections As ArrayList = New ArrayList()

        Public Sub SaveData()
            Dim xmlSerializer As XmlSerializer = New ActiveQueryBuilder.Core.Serialization.XmlSerializer()

            For Each connection As ConnectionInfo In _connections
                connection.ConnectionString = connection.ConnectionDescriptor.ConnectionString
                connection.LoadingOptions = xmlSerializer.Serialize(connection.ConnectionDescriptor.MetadataLoadingOptions)
                connection.SyntaxProviderState = xmlSerializer.SerializeObject(connection.ConnectionDescriptor.SyntaxProvider)
            Next
        End Sub

        Public Sub RemoveObsoleteConnectionInfos()
            Dim connectionsToRemove As List(Of ConnectionInfo) = New List(Of ConnectionInfo)()

            For Each connection As ConnectionInfo In _connections

                If connection.ConnectionDescriptor Is Nothing Then
                    connectionsToRemove.Add(connection)
                End If
            Next

            For Each connection As ConnectionInfo In connectionsToRemove
                _connections.Remove(connection)
            Next
        End Sub

        Public Sub RestoreData()
            Dim xmlSerializer As XmlSerializer = New XmlSerializer()

            For Each connection As ConnectionInfo In _connections
                If connection.ConnectionDescriptor Is Nothing Then Continue For
                connection.ConnectionDescriptor.ConnectionString = connection.ConnectionString

                If Not String.IsNullOrEmpty(connection.LoadingOptions) Then
                    xmlSerializer.Deserialize(connection.LoadingOptions, connection.ConnectionDescriptor.MetadataLoadingOptions)
                End If

                If Not String.IsNullOrEmpty(connection.SyntaxProviderName) AndAlso connection.IsGenericConnection() Then
                    connection.ConnectionDescriptor.SyntaxProvider = ConnectionInfo.GetSyntaxByName(connection.SyntaxProviderName)
                End If

                If Not String.IsNullOrEmpty(connection.SyntaxProviderState) Then
                    xmlSerializer.DeserializeObject(connection.SyntaxProviderState, connection.ConnectionDescriptor.SyntaxProvider)
                    connection.ConnectionDescriptor.RecreateSyntaxProperties()
                End If
            Next
        End Sub

        Default Public ReadOnly Property Item(index As Integer) As ConnectionInfo
            Get
                Return CType(_connections(index), ConnectionInfo)
            End Get
        End Property

        Public ReadOnly Property Count As Integer
            Get
                Return _connections.Count
            End Get
        End Property

        Public Property Connections As System.Collections.ArrayList
            Get
                Return _connections
            End Get
            Set(value As System.Collections.ArrayList)
                _connections = value
            End Set
        End Property

        Public Sub Add(ci As ConnectionInfo)
            _connections.Add(ci)
        End Sub

        Public Sub Insert(index As Integer, ci As ConnectionInfo)
            _connections.Insert(index, ci)
        End Sub

        Public Sub Remove(ci As ConnectionInfo)
            _connections.Remove(ci)
        End Sub
    End Class

    <Serializable>
    Public Class ConnectionInfo
        Public Property Name As String
        <XmlIgnore>
        Public Property ConnectionDescriptor As BaseConnectionDescriptor
        Public Property ConnectionString As String
        Public Property IsXmlFile As Boolean
        Public Property XMLPath As String
        Public Property CacheFile As String
        Public Property UserQueries As String
        Public Property MetadataStructure As String
        Public Property LoadingOptions As String
        Public Property SyntaxProviderState As String
        Public Property SyntaxProviderName As String

        Public Function IsGenericConnection() As Boolean
            Return TypeOf ConnectionDescriptor Is OLEDBConnectionDescriptor OrElse TypeOf ConnectionDescriptor Is ODBCConnectionDescriptor
        End Function

        Public Shared Function GetSyntaxByName(name As String) As BaseSyntaxProvider
            For Each syntax As Type In Helpers.SyntaxProviderList

                If syntax.ToString() = name Then
                    Return TryCast(Activator.CreateInstance(syntax), BaseSyntaxProvider)
                End If
            Next

            Return Nothing
        End Function

        Private _type As ConnectionTypes = ConnectionTypes.MSSQL

        Public Property Type As ConnectionTypes
            Get
                Return _type
            End Get
            Set(value As ConnectionTypes)
                _type = value
                CreateConnectionByType()

                If Not String.IsNullOrEmpty(SyntaxProviderName) AndAlso IsGenericConnection() Then
                    ConnectionDescriptor.SyntaxProvider = GetSyntaxByName(SyntaxProviderName)
                End If
            End Set
        End Property

        Public Sub New(descriptor As BaseConnectionDescriptor, name As String, type As ConnectionTypes, connectionString As String)
            Me.Name = name
            Me.ConnectionDescriptor = descriptor
            Me.Type = type
            Me.ConnectionString = connectionString
            Me.IsXmlFile = False
            Me.ConnectionDescriptor.ConnectionString = connectionString
        End Sub

        Public Sub New(xmlPath As String, name As String, type As ConnectionTypes)
            Me.Name = name
            Me.XMLPath = xmlPath
            Me.Type = type
            Me.IsXmlFile = True
            Me.ConnectionString = String.Empty
            Me.CreateConnectionByType()
        End Sub

        Public Sub New()
        End Sub

        Private Sub CreateConnectionByType()
            Try

                Select Case Type
                    Case ConnectionTypes.MSAccess
                        ConnectionDescriptor = New MSAccessConnectionDescriptor()
                        Return
                    Case ConnectionTypes.MSSQL
                        ConnectionDescriptor = New MSSQLConnectionDescriptor()
                        Return
                    Case ConnectionTypes.MSSQLAzure
                        ConnectionDescriptor = New MSSQLAzureConnectionDescriptor()
                        Return
                    Case ConnectionTypes.MySQL
                        ConnectionDescriptor = New MySQLConnectionDescriptor()
                        Return
                    Case ConnectionTypes.Oracle
                        ConnectionDescriptor = New OracleNativeConnectionDescriptor()
                        Return
                    Case ConnectionTypes.PostgreSQL
                        ConnectionDescriptor = New PostgreSQLConnectionDescriptor()
                        Return
                    Case ConnectionTypes.ODBC
                        ConnectionDescriptor = New ODBCConnectionDescriptor()
                        Return
                    Case ConnectionTypes.OLEDB
                        ConnectionDescriptor = New OLEDBConnectionDescriptor()
                        Return
                    Case ConnectionTypes.Firebird
                        ConnectionDescriptor = New FirebirdConnectionDescriptor()
                        Return
                    Case ConnectionTypes.SQLite
                        ConnectionDescriptor = New SQLiteConnectionDescriptor()
                        Return
                    Case ConnectionTypes.Excel
                        ConnectionDescriptor = New ExcelConnectionDescriptor()
                        Return
                End Select

            Catch
            End Try
        End Sub

        Public Function GetConnectionType(descriptorType As Type) As ConnectionTypes
            If descriptorType = GetType(MSAccessConnectionDescriptor) Then
                Return ConnectionTypes.MSAccess
            End If

            If descriptorType = GetType(ExcelConnectionDescriptor) Then
                Return ConnectionTypes.Excel
            End If

            If descriptorType = GetType(PostgreSQLConnectionDescriptor) Then
                Return ConnectionTypes.PostgreSQL
            End If

            If descriptorType = GetType(MSSQLConnectionDescriptor) Then
                Return ConnectionTypes.MSSQL
            End If

            If descriptorType = GetType(MSSQLAzureConnectionDescriptor) Then
                Return ConnectionTypes.MSSQLAzure
            End If

            If descriptorType = GetType(MySQLConnectionDescriptor) Then
                Return ConnectionTypes.MySQL
            End If

            If descriptorType = GetType(OracleNativeConnectionDescriptor) Then
                Return ConnectionTypes.Oracle
            End If

            If descriptorType = GetType(ODBCConnectionDescriptor) Then
                Return ConnectionTypes.ODBC
            End If

            If descriptorType = GetType(OLEDBConnectionDescriptor) Then
                Return ConnectionTypes.OLEDB
            End If

            If descriptorType = GetType(FirebirdConnectionDescriptor) Then
                Return ConnectionTypes.Firebird
            End If

            If descriptorType = GetType(SQLiteConnectionDescriptor) Then
                Return ConnectionTypes.SQLite
            End If

            Return ConnectionTypes.MSSQL
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If obj IsNot Nothing AndAlso TypeOf obj Is ConnectionInfo Then

                If (CType(obj, ConnectionInfo)).Type = Me.Type AndAlso (CType(obj, ConnectionInfo)).Name = Me.Name AndAlso (CType(obj, ConnectionInfo)).ConnectionString = Me.ConnectionString AndAlso (CType(obj, ConnectionInfo)).IsXmlFile = Me.IsXmlFile Then
                    Return True
                End If
            End If

            Return False
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return MyBase.GetHashCode()
        End Function
    End Class
