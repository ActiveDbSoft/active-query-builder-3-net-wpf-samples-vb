''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2021 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Xml.Serialization
Imports ActiveQueryBuilder.Core


Public Enum SourceType
    File
    [New]
    UserQueries
End Enum

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

<Serializable, XmlInclude(GetType(ConnectionInfo))>
Public Class ConnectionList
    <XmlElement(Type:=GetType(ConnectionInfo))>
    Private _connections As New ArrayList()

    Public Sub SaveData()
        Dim xmlSerializer = New ActiveQueryBuilder.Core.Serialization.XmlSerializer()
        For Each connection As ConnectionInfo In _connections
            connection.SyntaxProviderName = connection.ConnectionDescriptor.SyntaxProvider.GetType().ToString()
            connection.ConnectionString = connection.ConnectionDescriptor.ConnectionString
            connection.LoadingOptions = xmlSerializer.Serialize(connection.ConnectionDescriptor.MetadataLoadingOptions)
            connection.SyntaxProviderState = xmlSerializer.SerializeObject(connection.ConnectionDescriptor.SyntaxProvider)
        Next connection
    End Sub

    Public Sub RemoveObsoleteConnectionInfos()
        Dim connectionsToRemove = New List(Of ConnectionInfo)()

        For Each connection As ConnectionInfo In _connections
            If connection.ConnectionDescriptor Is Nothing Then
                connectionsToRemove.Add(connection)
            End If
        Next connection

        For Each connection As ConnectionInfo In connectionsToRemove
            _connections.Remove(connection)
        Next connection
    End Sub

    Public Sub RestoreData()
        Dim xmlSerializer = New ActiveQueryBuilder.Core.Serialization.XmlSerializer()

        For Each connection As ConnectionInfo In _connections
            If connection.ConnectionDescriptor Is Nothing Then
                Continue For
            End If
            Try
                connection.ConnectionDescriptor.ConnectionString = connection.ConnectionString

                If Not String.IsNullOrEmpty(connection.LoadingOptions) Then
                    xmlSerializer.Deserialize(connection.LoadingOptions, connection.ConnectionDescriptor.MetadataLoadingOptions)
                End If

                If Not String.IsNullOrEmpty(connection.SyntaxProviderName) AndAlso connection.IsGenericConnection() Then
                    connection.ConnectionDescriptor.SyntaxProvider = ConnectionInfo.GetSyntaxByName(connection.SyntaxProviderName)
                End If

                If Not String.IsNullOrEmpty(connection.SyntaxProviderState) Then

                    If Not String.IsNullOrEmpty(connection.SyntaxProviderName) Then
                        connection.ConnectionDescriptor.SyntaxProvider = ConnectionInfo.GetSyntaxByName(connection.SyntaxProviderName)
                    End If

                    xmlSerializer.DeserializeObject(connection.SyntaxProviderState, connection.ConnectionDescriptor.SyntaxProvider)
                    connection.ConnectionDescriptor.RecreateSyntaxProperties()
                End If
            Catch
                'ignore
            End Try

        Next connection
    End Sub

    Default Public ReadOnly Property Item(index As Integer) As ConnectionInfo
        Get
            Return DirectCast(_connections(index), ConnectionInfo)
        End Get
    End Property

    Public ReadOnly Property Count() As Integer
        Get
            Return _connections.Count
        End Get
    End Property

    Public Property Connections() As ArrayList
        Get
            Return _connections
        End Get
        Set(value As ArrayList)
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

    Public Property Name() As String
    Private _myConnectionDescriptor As BaseConnectionDescriptor
    Private _mySyntaxProvider As BaseSyntaxProvider

    <XmlIgnore>
    Public Property ConnectionDescriptor As BaseConnectionDescriptor

        Get
            Return _myConnectionDescriptor
        End Get
        Set
            _myConnectionDescriptor = Value
            If Not IsNothing(_myConnectionDescriptor) And Not IsNothing(_myConnectionDescriptor.MetadataProvider) And
               Not IsNothing(_myConnectionDescriptor.MetadataProvider.Connection) And
            String.IsNullOrEmpty(_myConnectionDescriptor.MetadataProvider.Connection.ConnectionString) And
            Not String.IsNullOrEmpty(ConnectionString) And Not IsXmlFile Then
                _myConnectionDescriptor.MetadataProvider.Connection.ConnectionString = ConnectionString
            End If

        End Set
    End Property

    <XmlIgnore>
    Public Property SyntaxProvider As BaseSyntaxProvider
        Get
            If Not IsNothing(_myConnectionDescriptor) And Not IsNothing(_myConnectionDescriptor.SyntaxProvider) Then
                Return _myConnectionDescriptor.SyntaxProvider
            Else
                Return _mySyntaxProvider
            End If

        End Get
        Set
            _mySyntaxProvider = Value
            If Not IsNothing(_myConnectionDescriptor) Then _myConnectionDescriptor.SyntaxProvider = Value
        End Set
    End Property

    Public Property ConnectionString() As String
    Public Property IsXmlFile() As Boolean
    Public Property XMLPath() As String
    Public Property UserQueries() As String
    Public Property LoadingOptions() As String
    Public Property SyntaxProviderState() As String
    Public Property SyntaxProviderName() As String

    Public Function IsGenericConnection() As Boolean
        Return TypeOf ConnectionDescriptor Is OLEDBConnectionDescriptor OrElse TypeOf ConnectionDescriptor Is ODBCConnectionDescriptor
    End Function

    Public Shared Function GetSyntaxByName(name As String) As BaseSyntaxProvider
        For Each syntax As Type In Helpers.SyntaxProviderList
            If syntax.ToString() = name Then
                Return TryCast(Activator.CreateInstance(syntax), BaseSyntaxProvider)
            End If
        Next syntax

        Return Nothing
    End Function

    Private _type As ConnectionTypes = ConnectionTypes.MSSQL

    Public Property Type() As ConnectionTypes
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
        ConnectionDescriptor = descriptor
        Me.Type = type
        Me.ConnectionString = connectionString
        IsXmlFile = False
        ConnectionDescriptor.ConnectionString = connectionString
    End Sub

    Public Sub New(xmlPath As String, name As String, type As ConnectionTypes)
        Me.Name = name
        Me.XMLPath = xmlPath
        Me.Type = type
        IsXmlFile = True
        ConnectionString = String.Empty
        CreateConnectionByType()
    End Sub

    Public Sub New(connectionType As ConnectionTypes, connectionName As String, connectionString As String, isFromXml As Boolean, userQueriesXml As String)
        Type = connectionType
        Name = connectionName
        Me.ConnectionString = connectionString
        IsXmlFile = isFromXml
        UserQueries = userQueriesXml
        CreateConnectionByType()
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
        Finally
            'ignore
        End Try
    End Sub

    Public Function GetConnectionType(descriptorType As Type) As ConnectionTypes
        If descriptorType Is GetType(MSAccessConnectionDescriptor) Then
            Return ConnectionTypes.MSAccess
        End If
        If descriptorType Is GetType(ExcelConnectionDescriptor) Then
            Return ConnectionTypes.Excel
        End If
        If descriptorType Is GetType(PostgreSQLConnectionDescriptor) Then
            Return ConnectionTypes.PostgreSQL
        End If
        If descriptorType Is GetType(MSSQLConnectionDescriptor) Then
            Return ConnectionTypes.MSSQL
        End If
        If descriptorType Is GetType(MSSQLAzureConnectionDescriptor) Then
            Return ConnectionTypes.MSSQLAzure
        End If
        If descriptorType Is GetType(MySQLConnectionDescriptor) Then
            Return ConnectionTypes.MySQL
        End If
        If descriptorType Is GetType(OracleNativeConnectionDescriptor) Then
            Return ConnectionTypes.Oracle
        End If
        If descriptorType Is GetType(ODBCConnectionDescriptor) Then
            Return ConnectionTypes.ODBC
        End If
        If descriptorType Is GetType(OLEDBConnectionDescriptor) Then
            Return ConnectionTypes.OLEDB
        End If
        If descriptorType Is GetType(FirebirdConnectionDescriptor) Then
            Return ConnectionTypes.Firebird
        End If
        If descriptorType Is GetType(SQLiteConnectionDescriptor) Then
            Return ConnectionTypes.SQLite
        End If

        Return ConnectionTypes.MSSQL
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        If Not (TypeOf obj Is ConnectionInfo) Then
            Return False
        End If

        Return DirectCast(obj, ConnectionInfo).Type = Type AndAlso DirectCast(obj, ConnectionInfo).Name = Name AndAlso DirectCast(obj, ConnectionInfo).ConnectionString = ConnectionString AndAlso DirectCast(obj, ConnectionInfo).IsXmlFile = IsXmlFile
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return MyBase.GetHashCode()
    End Function
End Class

Public Class Misc
    Public Shared ReadOnly ConnectionDescriptorList As New List(Of Type) From {GetType(MSAccessConnectionDescriptor), GetType(ExcelConnectionDescriptor), GetType(MSSQLConnectionDescriptor), GetType(MSSQLAzureConnectionDescriptor), GetType(MySQLConnectionDescriptor), GetType(OracleNativeConnectionDescriptor), GetType(PostgreSQLConnectionDescriptor), GetType(ODBCConnectionDescriptor), GetType(OLEDBConnectionDescriptor), GetType(SQLiteConnectionDescriptor), GetType(FirebirdConnectionDescriptor)}

    Public Shared ReadOnly ConnectionDescriptorNames As New List(Of String) From {"MS Access", "Excel", "MS SQL Server", "MS SQL Server Azure", "MySQL", "Oracle Native", "PostgreSQL", "Generic ODBC Connection", "Generic OLEDB Connection", "SQLite", "Firebird"}
End Class
