'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Collections
Imports System.Configuration
Imports System.Xml.Serialization
Imports ActiveQueryBuilder.Core

Public Enum ConnectionTypes
    MSSQL
    MSAccess
    Oracle
    MySQL
    PostgreSQL
    OLEDB
    ODBC
End Enum

<Serializable>
<XmlInclude(GetType(ConnectionInfo))>
<SettingsSerializeAs(SettingsSerializeAs.[String])>
Public Class ConnectionList
    <XmlElement(Type:=GetType(ConnectionInfo))>
    Private _connections As New ArrayList()

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
        Set
            _connections = Value
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
    Private _syntaxProvider As BaseSyntaxProvider
    Private _syntaxProviderName As String

    Public ConnectionType As ConnectionTypes
    Public ConnectionName As String
    Public ConnectionString As String
    Public IsXmlFile As Boolean
    Public CacheFile As String
    Public UserQueries As String

    

    Public Property SyntaxProviderName As String
        Set
            If _syntaxProviderName = value Then Return
            _syntaxProviderName = value
            Dim foundSyntaxProviderType As Type = GetType(GenericSyntaxProvider)
            For Each syntaxProviderType As Type In Helpers.SyntaxProviderList
                If String.Equals(syntaxProviderType.Name, value, StringComparison.InvariantCultureIgnoreCase) Then
                    foundSyntaxProviderType = syntaxProviderType
                    Exit For
                End If
            Next

            For Each syntaxProviderType As Type In Helpers.SyntaxProviderList
                Using syntaxProvider1 As BaseSyntaxProvider = CType(Activator.CreateInstance(syntaxProviderType), BaseSyntaxProvider)
                    If String.Equals(syntaxProvider1.Description, value, StringComparison.InvariantCultureIgnoreCase) Then
                        foundSyntaxProviderType = syntaxProviderType
                        Exit For
                    End If
                End Using
            Next

            If _syntaxProvider IsNot Nothing AndAlso _syntaxProvider.[GetType]() = foundSyntaxProviderType Then Return
            If _syntaxProvider IsNot Nothing Then _syntaxProvider.Dispose()
            _syntaxProvider = CType(Activator.CreateInstance(foundSyntaxProviderType), BaseSyntaxProvider)
        End Set

        Get
            Return _syntaxProviderName
        End Get
    End Property

    <XmlIgnore>
    Public Property SyntaxProvider As BaseSyntaxProvider
        Set
            _syntaxProvider = value
            If _syntaxProvider IsNot Nothing Then SyntaxProviderName = _syntaxProvider.[GetType]().Name
        End Set

        Get
            Return _syntaxProvider
        End Get
    End Property

    Public Sub New()
        ConnectionType = ConnectionTypes.MSSQL

        ConnectionName = Nothing
        ConnectionString = Nothing
        IsXmlFile = False
        CacheFile = Nothing
    End Sub

    Public Sub New(connectionType1 As ConnectionTypes, connectionName1 As String, connectionString1 As String, isFromXml As Boolean, cacheFile1 As String, userQueriesXml As String)
        ConnectionType = connectionType1
        ConnectionName = connectionName1
        ConnectionString = connectionString1
        IsXmlFile = isFromXml
        CacheFile = cacheFile1
        UserQueries = userQueriesXml
    End Sub

    Public Overrides Function Equals(obj As Object) As Boolean
        If TypeOf obj Is ConnectionInfo Then
            If DirectCast(obj, ConnectionInfo).ConnectionType = ConnectionType AndAlso DirectCast(obj, ConnectionInfo).ConnectionName = ConnectionName AndAlso DirectCast(obj, ConnectionInfo).ConnectionString = ConnectionString AndAlso DirectCast(obj, ConnectionInfo).IsXmlFile = IsXmlFile Then
                Return True
            End If
        End If

        Return False
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return MyBase.GetHashCode()
    End Function
End Class
