'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Threading
Imports ActiveQueryBuilder.Core
Imports FullFeaturedDemo.Connection.FrameConnection

Namespace Connection
    ''' <summary>
    ''' Interaction logic for AddConnectionWindow.xaml
    ''' </summary>
    Partial Public Class AddConnectionWindow
        Private ReadOnly _connectionInfo As ConnectionInfo
        Private _currentConnectionFrame As IConnectionFrame

        Public Sub New()
            InitializeComponent()
            AddHandler Dispatcher.CurrentDispatcher.Hooks.DispatcherInactive, AddressOf Hooks_DispatcherInactive
        End Sub

        Private Sub Hooks_DispatcherInactive(sender As Object, e As EventArgs)
            ButtonOk.IsEnabled = (TextBoxConnectionName.Text.Length > 0)
        End Sub

        Protected Overrides Sub OnClosing(e As CancelEventArgs)
            If DialogResult = True Then
                If _currentConnectionFrame IsNot Nothing AndAlso _currentConnectionFrame.TestConnection() Then
                    _connectionInfo.ConnectionName = TextBoxConnectionName.Text
                    _connectionInfo.ConnectionString = _currentConnectionFrame.ConnectionString
                    e.Cancel = False
                    RemoveHandler Dispatcher.CurrentDispatcher.Hooks.DispatcherInactive, AddressOf Hooks_DispatcherInactive
                Else
                    e.Cancel = True
                End If
            End If

            MyBase.OnClosing(e)
        End Sub

        Public Sub New(connectionInfo As ConnectionInfo)
            InitializeComponent()

            Debug.Assert(connectionInfo IsNot Nothing)

            _connectionInfo = connectionInfo
            TextBoxConnectionName.Text = connectionInfo.ConnectionName

            If Not String.IsNullOrEmpty(connectionInfo.ConnectionString) Then

                If Not connectionInfo.IsXmlFile Then
                    rbMSSQL.IsEnabled = (connectionInfo.ConnectionType = ConnectionTypes.MSSQL)
                    rbMSAccess.IsEnabled = (connectionInfo.ConnectionType = ConnectionTypes.MSAccess)
                    rbOracle.IsEnabled = (connectionInfo.ConnectionType = ConnectionTypes.Oracle)
                    rbMySQL.IsEnabled = (connectionInfo.ConnectionType = ConnectionTypes.MySQL)
                    rbPostrgeSQL.IsEnabled = (connectionInfo.ConnectionType = ConnectionTypes.PostgreSQL)
                    rbOLEDB.IsEnabled = (connectionInfo.ConnectionType = ConnectionTypes.OLEDB)
                    rbODBC.IsEnabled = (connectionInfo.ConnectionType = ConnectionTypes.ODBC)
                End If
            End If

            If connectionInfo.IsXmlFile Then
                rbOLEDB.IsEnabled = False
                rbODBC.IsEnabled = False
            End If

            rbMSSQL.IsChecked = (connectionInfo.ConnectionType = ConnectionTypes.MSSQL)
            rbMSAccess.IsChecked = (connectionInfo.ConnectionType = ConnectionTypes.MSAccess)
            rbOracle.IsChecked = (connectionInfo.ConnectionType = ConnectionTypes.Oracle)
            rbMySQL.IsChecked = (connectionInfo.ConnectionType = ConnectionTypes.MySQL)
            rbPostrgeSQL.IsChecked = (connectionInfo.ConnectionType = ConnectionTypes.PostgreSQL)
            rbOLEDB.IsChecked = (connectionInfo.ConnectionType = ConnectionTypes.OLEDB)
            rbODBC.IsChecked = (connectionInfo.ConnectionType = ConnectionTypes.ODBC)

            SetActiveConnectionTypeFrame()

            AddHandler rbMSSQL.Checked, AddressOf ConnectionTypeChanged
            AddHandler rbMSAccess.Checked, AddressOf ConnectionTypeChanged
            AddHandler rbOracle.Checked, AddressOf ConnectionTypeChanged
            AddHandler rbMySQL.Checked, AddressOf ConnectionTypeChanged
            AddHandler rbPostrgeSQL.Checked, AddressOf ConnectionTypeChanged
            AddHandler rbOLEDB.Checked, AddressOf ConnectionTypeChanged
            AddHandler rbODBC.Checked, AddressOf ConnectionTypeChanged

            FillSyntax()
        End Sub

        Private Sub FillSyntax()
            BoxSyntaxProvider.Items.Clear()
            BoxServerVersion.Items.Clear()

            If Not String.IsNullOrEmpty(_connectionInfo.SyntaxProviderName) AndAlso _connectionInfo.SyntaxProvider Is Nothing Then
                BoxSyntaxProvider.Items.Add(_connectionInfo.SyntaxProviderName)
                BoxSyntaxProvider.SelectedItem = _connectionInfo.SyntaxProviderName
                Return
            End If

            If _connectionInfo.SyntaxProvider Is Nothing Then
                Select Case _connectionInfo.ConnectionType
                    Case ConnectionTypes.MSSQL
                        _connectionInfo.SyntaxProvider = New MSSQLSyntaxProvider()
                        Exit Select
                    Case ConnectionTypes.MSAccess
                        _connectionInfo.SyntaxProvider = New MSAccessSyntaxProvider()
                        Exit Select
                    Case ConnectionTypes.Oracle
                        _connectionInfo.SyntaxProvider = New OracleSyntaxProvider()
                        Exit Select
                    Case ConnectionTypes.MySQL
                        _connectionInfo.SyntaxProvider = New MySQLSyntaxProvider()
                        Exit Select
                    Case ConnectionTypes.PostgreSQL
                        _connectionInfo.SyntaxProvider = New PostgreSQLSyntaxProvider()
                        Exit Select
                    Case ConnectionTypes.OLEDB
                        _connectionInfo.SyntaxProvider = New SQL92SyntaxProvider()
                        Exit Select
                    Case ConnectionTypes.ODBC
                        _connectionInfo.SyntaxProvider = New SQL92SyntaxProvider()
                        Exit Select
                End Select
            End If

            If TypeOf _connectionInfo.SyntaxProvider Is SQL2003SyntaxProvider Then
                BoxSyntaxProvider.Items.Add("ANSI SQL-2003")
                BoxSyntaxProvider.SelectedItem = "ANSI SQL-2003"

                BoxSyntaxProvider.Items.Add("ANSI SQL-92")
                BoxSyntaxProvider.Items.Add("ANSI SQL-89")
                BoxSyntaxProvider.Items.Add("Firebird")
                BoxSyntaxProvider.Items.Add("IBM DB2")
                BoxSyntaxProvider.Items.Add("IBM Informix")
                BoxSyntaxProvider.Items.Add("Microsoft Access")
                BoxSyntaxProvider.Items.Add("Microsoft SQL Server")
                BoxSyntaxProvider.Items.Add("MySQL")
                BoxSyntaxProvider.Items.Add("Oracle")
                BoxSyntaxProvider.Items.Add("PostgreSQL")
                BoxSyntaxProvider.Items.Add("SQLite")
                BoxSyntaxProvider.Items.Add("Sybase")
                BoxSyntaxProvider.Items.Add("VistaDB")
                BoxSyntaxProvider.Items.Add("Universal")
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is SQL92SyntaxProvider Then
                BoxSyntaxProvider.Items.Add("ANSI SQL-92")
                BoxSyntaxProvider.SelectedItem = "ANSI SQL-92"

                BoxSyntaxProvider.Items.Add("ANSI SQL-2003")

                BoxSyntaxProvider.Items.Add("ANSI SQL-89")
                BoxSyntaxProvider.Items.Add("Firebird")
                BoxSyntaxProvider.Items.Add("IBM DB2")
                BoxSyntaxProvider.Items.Add("IBM Informix")
                BoxSyntaxProvider.Items.Add("Microsoft Access")
                BoxSyntaxProvider.Items.Add("Microsoft SQL Server")
                BoxSyntaxProvider.Items.Add("MySQL")
                BoxSyntaxProvider.Items.Add("Oracle")
                BoxSyntaxProvider.Items.Add("PostgreSQL")
                BoxSyntaxProvider.Items.Add("SQLite")
                BoxSyntaxProvider.Items.Add("Sybase")
                BoxSyntaxProvider.Items.Add("VistaDB")
                BoxSyntaxProvider.Items.Add("Universal")
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is SQL89SyntaxProvider Then
                BoxSyntaxProvider.Items.Add("ANSI SQL-89")
                BoxSyntaxProvider.SelectedItem = "ANSI SQL-89"

                BoxSyntaxProvider.Items.Add("ANSI SQL-2003")
                BoxSyntaxProvider.Items.Add("ANSI SQL-92")
                BoxSyntaxProvider.Items.Add("Firebird")
                BoxSyntaxProvider.Items.Add("IBM DB2")
                BoxSyntaxProvider.Items.Add("IBM Informix")
                BoxSyntaxProvider.Items.Add("Microsoft Access")
                BoxSyntaxProvider.Items.Add("Microsoft SQL Server")
                BoxSyntaxProvider.Items.Add("MySQL")
                BoxSyntaxProvider.Items.Add("Oracle")
                BoxSyntaxProvider.Items.Add("PostgreSQL")
                BoxSyntaxProvider.Items.Add("SQLite")
                BoxSyntaxProvider.Items.Add("Sybase")
                BoxSyntaxProvider.Items.Add("VistaDB")
                BoxSyntaxProvider.Items.Add("Universal")
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is FirebirdSyntaxProvider Then
                BoxSyntaxProvider.Items.Add("Firebird")
                BoxSyntaxProvider.SelectedItem = "Firebird"
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is DB2SyntaxProvider Then
                BoxSyntaxProvider.Items.Add("IBM DB2")
                BoxSyntaxProvider.SelectedItem = "IBM DB2"
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is InformixSyntaxProvider Then
                BoxSyntaxProvider.Items.Add("IBM Informix")
                BoxSyntaxProvider.SelectedItem = "IBM Informix"
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is MSAccessSyntaxProvider Then
                BoxSyntaxProvider.Items.Add("Microsoft Access")
                BoxSyntaxProvider.SelectedItem = "Microsoft Access"
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is MSSQLSyntaxProvider Then
                BoxSyntaxProvider.Items.Add("Microsoft SQL Server")
                BoxSyntaxProvider.SelectedItem = "Microsoft SQL Server"
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is MySQLSyntaxProvider Then
                BoxSyntaxProvider.Items.Add("MySQL")
                BoxSyntaxProvider.SelectedItem = "MySQL"
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is OracleSyntaxProvider Then
                BoxSyntaxProvider.Items.Add("Oracle")
                BoxSyntaxProvider.SelectedItem = "Oracle"
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is PostgreSQLSyntaxProvider Then
                BoxSyntaxProvider.Items.Add("PostgreSQL")
                BoxSyntaxProvider.SelectedItem = "PostgreSQL"
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is SQLiteSyntaxProvider Then
                BoxSyntaxProvider.Items.Add("SQLite")
                BoxSyntaxProvider.SelectedItem = "SQLite"
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is SybaseSyntaxProvider Then
                BoxSyntaxProvider.Items.Add("Sybase")
                BoxSyntaxProvider.SelectedItem = "Sybase"
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is VistaDBSyntaxProvider Then
                BoxSyntaxProvider.Items.Add("VistaDB")
                BoxSyntaxProvider.SelectedItem = "VistaDB"
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is GenericSyntaxProvider Then
                BoxSyntaxProvider.Items.Add("Universal")
                BoxSyntaxProvider.SelectedItem = "Universal"


                BoxSyntaxProvider.Items.Add("ANSI SQL-2003")
                BoxSyntaxProvider.Items.Add("ANSI SQL-92")
                BoxSyntaxProvider.Items.Add("ANSI SQL-89")
                BoxSyntaxProvider.Items.Add("Firebird")
                BoxSyntaxProvider.Items.Add("IBM DB2")
                BoxSyntaxProvider.Items.Add("IBM Informix")
                BoxSyntaxProvider.Items.Add("Microsoft Access")
                BoxSyntaxProvider.Items.Add("Microsoft SQL Server")
                BoxSyntaxProvider.Items.Add("MySQL")
                BoxSyntaxProvider.Items.Add("Oracle")
                BoxSyntaxProvider.Items.Add("PostgreSQL")
                BoxSyntaxProvider.Items.Add("SQLite")
                BoxSyntaxProvider.Items.Add("Sybase")
                BoxSyntaxProvider.Items.Add("VistaDB")
                BoxSyntaxProvider.Items.Add("Universal")
            End If


            FillVersions()
        End Sub

        Private Sub FillVersions()
            BoxServerVersion.Items.Clear()
            BoxServerVersion.Text = String.Empty

            If TypeOf _connectionInfo.SyntaxProvider Is SQL2003SyntaxProvider Then
                BoxServerVersion.IsEnabled = False
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is SQL92SyntaxProvider Then
                BoxServerVersion.IsEnabled = False
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is SQL89SyntaxProvider Then
                BoxServerVersion.IsEnabled = False
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is FirebirdSyntaxProvider Then
                BoxServerVersion.IsEnabled = True
                BoxServerVersion.Items.Add("Firebird 1.0")
                BoxServerVersion.Items.Add("Firebird 1.5")
                BoxServerVersion.Items.Add("Firebird 2.0")
                BoxServerVersion.Items.Add("Firebird 2.5")

                Dim firebirdSyntaxProvider As FirebirdSyntaxProvider = CType(_connectionInfo.SyntaxProvider, FirebirdSyntaxProvider)

                Select Case firebirdSyntaxProvider.ServerVersion
                    Case FirebirdVersion.Firebird10
                        BoxServerVersion.SelectedItem = "Firebird 1.0"
                        Exit Sub
                    Case FirebirdVersion.Firebird15
                        BoxServerVersion.SelectedItem = "Firebird 1.5"
                        Exit Sub
                    Case FirebirdVersion.Firebird20
                        BoxServerVersion.SelectedItem = "Firebird 2.0"
                        Exit Sub
                    Case FirebirdVersion.Firebird25
                        BoxServerVersion.SelectedItem = "Firebird 2.5"
                        Exit Sub
                End Select
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is DB2SyntaxProvider Then
                BoxServerVersion.IsEnabled = False
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is InformixSyntaxProvider Then
                BoxServerVersion.IsEnabled = True
                BoxServerVersion.Items.Add("Informix 8")
                BoxServerVersion.Items.Add("Informix 9")
                BoxServerVersion.Items.Add("Informix 10")
                BoxServerVersion.Items.Add("Informix 11")

                Dim informixSyntaxProvider As InformixSyntaxProvider = CType(_connectionInfo.SyntaxProvider, InformixSyntaxProvider)

                Select Case informixSyntaxProvider.ServerVersion
                    Case InformixVersion.DS8
                        BoxServerVersion.SelectedItem = "Informix 8"
                        Exit Sub
                    Case InformixVersion.DS9
                        BoxServerVersion.SelectedItem = "Informix 9"
                        Exit Sub
                    Case InformixVersion.DS10
                        BoxServerVersion.SelectedItem = "Informix 10"
                        Exit Sub
                    Case Else
                        BoxServerVersion.SelectedItem = "Informix 11"
                        Exit Sub
                End Select
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is MSAccessSyntaxProvider Then
                BoxServerVersion.IsEnabled = True
                BoxServerVersion.Items.Add("MS Jet 3")
                BoxServerVersion.Items.Add("MS Jet 4")

                Dim accessSyntaxProvider As MSAccessSyntaxProvider = CType(_connectionInfo.SyntaxProvider, MSAccessSyntaxProvider)

                If (Equals(accessSyntaxProvider.ServerVersion, MSAccessServerVersion.MSJET3)) Then
                    BoxServerVersion.SelectedItem = "MS Jet 3"
                Else
                    BoxServerVersion.SelectedItem = "MS Jet 4"
                End If


            ElseIf TypeOf _connectionInfo.SyntaxProvider Is MSSQLSyntaxProvider Then
                BoxServerVersion.IsEnabled = True
                BoxServerVersion.Items.Add("Auto")
                BoxServerVersion.Items.Add("SQL Server 7")
                BoxServerVersion.Items.Add("SQL Server 2000")
                BoxServerVersion.Items.Add("SQL Server 2005")
                BoxServerVersion.Items.Add("SQL Server 2012")
                BoxServerVersion.Items.Add("SQL Server 2014")

                Dim mssqlSyntaxProvider As MSSQLSyntaxProvider = CType(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider)

                Select Case mssqlSyntaxProvider.ServerVersion
                    Case MSSQLServerVersion.MSSQL7
                        BoxServerVersion.SelectedItem = "SQL Server 7"
                        Exit Sub
                    Case MSSQLServerVersion.MSSQL2000
                        BoxServerVersion.SelectedItem = "SQL Server 2000"
                        Exit Sub
                    Case MSSQLServerVersion.MSSQL2005
                        BoxServerVersion.SelectedItem = "SQL Server 2005"
                        Exit Sub
                    Case MSSQLServerVersion.MSSQL2012
                        BoxServerVersion.SelectedItem = "SQL Server 2012"
                        Exit Sub
                    Case MSSQLServerVersion.MSSQL2014
                        BoxServerVersion.SelectedItem = "SQL Server 2014"
                        Exit Sub
                    Case Else
                        BoxServerVersion.SelectedItem = "Auto"
                        Exit Sub
                End Select
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is MySQLSyntaxProvider Then
                BoxServerVersion.IsEnabled = True
                BoxServerVersion.Items.Add("3.0")
                BoxServerVersion.Items.Add("4.0")
                BoxServerVersion.Items.Add("5.0+")

                Dim mySqlSyntaxProvider As MySQLSyntaxProvider = CType(_connectionInfo.SyntaxProvider, MySQLSyntaxProvider)

                If mySqlSyntaxProvider.ServerVersionInt < 40000 Then
                    BoxServerVersion.SelectedItem = "3.0"
                ElseIf mySqlSyntaxProvider.ServerVersionInt < 50000 Then
                    BoxServerVersion.SelectedItem = "4.0"
                Else
                    BoxServerVersion.SelectedItem = "5.0+"
                End If
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is OracleSyntaxProvider Then
                BoxServerVersion.IsEnabled = True
                BoxServerVersion.Items.Add("Oracle 7")
                BoxServerVersion.Items.Add("Oracle 8")
                BoxServerVersion.Items.Add("Oracle 9")
                BoxServerVersion.Items.Add("Oracle 10")
                BoxServerVersion.Items.Add("Oracle 11")
                BoxServerVersion.Items.Add("Oracle 12")

                Dim oracleSyntaxProvider As OracleSyntaxProvider = CType(_connectionInfo.SyntaxProvider, OracleSyntaxProvider)

                Select Case oracleSyntaxProvider.ServerVersion
                    Case OracleServerVersion.Oracle7
                        BoxServerVersion.SelectedItem = "Oracle 7"
                        Exit Sub
                    Case OracleServerVersion.Oracle8
                        BoxServerVersion.SelectedItem = "Oracle 8"
                        Exit Sub
                    Case OracleServerVersion.Oracle9
                        BoxServerVersion.SelectedItem = "Oracle 9"
                        Exit Sub
                    Case OracleServerVersion.Oracle10
                        BoxServerVersion.SelectedItem = "Oracle 10"
                        Exit Sub
                    Case OracleServerVersion.Oracle11
                        BoxServerVersion.SelectedItem = "Oracle 11"
                        Exit Sub
                    Case Else
                        BoxServerVersion.SelectedItem = "Oracle 12"
                        Exit Sub
                End Select
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is PostgreSQLSyntaxProvider Then
                BoxServerVersion.IsEnabled = False
                BoxServerVersion.Text = "Auto"
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is SQLiteSyntaxProvider Then
                BoxServerVersion.IsEnabled = False
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is SybaseSyntaxProvider Then
                BoxServerVersion.IsEnabled = True
                BoxServerVersion.Items.Add("ASE")
                BoxServerVersion.Items.Add("SQL Anywhere")
                BoxServerVersion.Items.Add("SAP IQ")

                Dim sybaseSyntaxProvider As SybaseSyntaxProvider = CType(_connectionInfo.SyntaxProvider, SybaseSyntaxProvider)

                Select Case sybaseSyntaxProvider.ServerVersion
                    Case SybaseServerVersion.SybaseASE
                        BoxServerVersion.SelectedItem = "ASE"
                        Exit Sub
                    Case SybaseServerVersion.SybaseASA
                        BoxServerVersion.SelectedItem = "SQL Anywhere"
                        Exit Sub
                    Case SybaseServerVersion.SybaseIQ
                        BoxServerVersion.SelectedItem = "SAP IQ"
                        Exit Sub
                End Select
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is VistaDBSyntaxProvider Then
                BoxServerVersion.IsEnabled = False
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is GenericSyntaxProvider Then
                BoxServerVersion.IsEnabled = False
            End If
        End Sub

        Private Sub ConnectionTypeChanged(sender As Object, e As RoutedEventArgs)
            If DirectCast(sender, RadioButton).IsChecked <> True Then
                Return
            End If

            Dim connectionType As ConnectionTypes = ConnectionTypes.MSSQL

            If Equals(sender, rbMSSQL) Then
                connectionType = ConnectionTypes.MSSQL
            ElseIf Equals(sender, rbMSAccess) Then
                connectionType = ConnectionTypes.MSAccess
            ElseIf Equals(sender, rbOracle) Then
                connectionType = ConnectionTypes.Oracle
            ElseIf Equals(sender, rbMySQL) Then
                connectionType = ConnectionTypes.MySQL
            ElseIf Equals(sender, rbPostrgeSQL) Then
                connectionType = ConnectionTypes.PostgreSQL
            ElseIf Equals(sender, rbOLEDB) Then
                connectionType = ConnectionTypes.OLEDB
            ElseIf Equals(sender, rbODBC) Then
                connectionType = ConnectionTypes.ODBC
            End If

            If connectionType <> _connectionInfo.ConnectionType Then
                _connectionInfo.ConnectionType = connectionType

                If Not _connectionInfo.IsXmlFile Then
                    SetActiveConnectionTypeFrame()
                End If
            End If

            Select Case _connectionInfo.ConnectionType
                Case ConnectionTypes.MSSQL
                    _connectionInfo.SyntaxProvider = New MSSQLSyntaxProvider()
                    Exit Select
                Case ConnectionTypes.MSAccess
                    _connectionInfo.SyntaxProvider = New MSAccessSyntaxProvider()
                    Exit Select
                Case ConnectionTypes.Oracle
                    _connectionInfo.SyntaxProvider = New OracleSyntaxProvider()
                    Exit Select
                Case ConnectionTypes.MySQL
                    _connectionInfo.SyntaxProvider = New MySQLSyntaxProvider()
                    Exit Select
                Case ConnectionTypes.PostgreSQL
                    _connectionInfo.SyntaxProvider = New PostgreSQLSyntaxProvider()
                    Exit Select
                Case ConnectionTypes.OLEDB
                    _connectionInfo.SyntaxProvider = New SQL92SyntaxProvider()
                    Exit Select
                Case ConnectionTypes.ODBC
                    _connectionInfo.SyntaxProvider = New SQL92SyntaxProvider()
                    Exit Select
            End Select

            FillSyntax()
        End Sub

        Private Sub SetActiveConnectionTypeFrame()
            If GridFrames.Children.Contains(DirectCast(_currentConnectionFrame, FrameworkElement)) Then
                GridFrames.Children.Remove(DirectCast(_currentConnectionFrame, FrameworkElement))
                _currentConnectionFrame = Nothing
            End If

            If Not _connectionInfo.IsXmlFile Then
                Select Case _connectionInfo.ConnectionType
                    Case ConnectionTypes.MSSQL
                        _currentConnectionFrame = New MSSQLConnectionFrame(_connectionInfo.ConnectionString)
                        Exit Select
                    Case ConnectionTypes.MSAccess
                        _currentConnectionFrame = New MSAccessConnectionFrame(_connectionInfo.ConnectionString)
                        Exit Select
                    Case ConnectionTypes.Oracle
                        _currentConnectionFrame = New OracleConnectionFrame(_connectionInfo.ConnectionString)
                        Exit Select
                    Case ConnectionTypes.MySQL
                        _currentConnectionFrame = New MySQLConnectionFrame(_connectionInfo.ConnectionString)
                        Exit Select
                    Case ConnectionTypes.PostgreSQL
                        _currentConnectionFrame = New PostgreSQLConnectionFrame(_connectionInfo.ConnectionString)
                        Exit Select
                    Case ConnectionTypes.OLEDB
                        _currentConnectionFrame = New OLEDBConnectionFrame(_connectionInfo.ConnectionString)
                        Exit Select
                    Case ConnectionTypes.ODBC
                        _currentConnectionFrame = New ODBCConnectionFrame(_connectionInfo.ConnectionString)
                        Exit Select
                End Select
            Else
                _currentConnectionFrame = New XmlFileFrame(_connectionInfo.ConnectionString)
            End If

            If _currentConnectionFrame IsNot Nothing Then
                GridFrames.Children.Add(DirectCast(_currentConnectionFrame, FrameworkElement))
            End If
        End Sub

        Private Sub ButtonBaseOK_OnClick(sender As Object, e As RoutedEventArgs)
            DialogResult = True
        End Sub

        Private Sub ButtonBaseClose_OnClick(sender As Object, e As RoutedEventArgs)
            Close()
        End Sub

        Private Sub BoxSyntaxProvider_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            Select Case DirectCast(BoxSyntaxProvider.SelectedItem, String)
                Case "ANSI SQL-2003"
                    _connectionInfo.SyntaxProvider = New SQL2003SyntaxProvider()
                    Exit Select
                Case "ANSI SQL-92"
                    _connectionInfo.SyntaxProvider = New SQL92SyntaxProvider()
                    Exit Select
                Case "ANSI SQL-89"
                    _connectionInfo.SyntaxProvider = New SQL89SyntaxProvider()
                    Exit Select
                Case "Firebird"
                    _connectionInfo.SyntaxProvider = New FirebirdSyntaxProvider()
                    Exit Select
                Case "IBM DB2"
                    _connectionInfo.SyntaxProvider = New DB2SyntaxProvider()
                    Exit Select
                Case "IBM Informix"
                    _connectionInfo.SyntaxProvider = New InformixSyntaxProvider()
                    Exit Select
                Case "Microsoft Access"
                    _connectionInfo.SyntaxProvider = New MSAccessSyntaxProvider()
                    Exit Select
                Case "Microsoft SQL Server"
                    _connectionInfo.SyntaxProvider = New MSSQLSyntaxProvider()
                    Exit Select
                Case "MySQL"
                    _connectionInfo.SyntaxProvider = New MySQLSyntaxProvider()
                    Exit Select
                Case "Oracle"
                    _connectionInfo.SyntaxProvider = New OracleSyntaxProvider()
                    Exit Select
                Case "PostgreSQL"
                    _connectionInfo.SyntaxProvider = New PostgreSQLSyntaxProvider()
                    Exit Select
                Case "SQLite"
                    _connectionInfo.SyntaxProvider = New SQLiteSyntaxProvider()
                    Exit Select
                Case "Sybase"
                    _connectionInfo.SyntaxProvider = New SybaseSyntaxProvider()
                    Exit Select
                Case "VistaDB"
                    _connectionInfo.SyntaxProvider = New VistaDBSyntaxProvider()
                    Exit Select
                Case "Universal"
                    _connectionInfo.SyntaxProvider = New GenericSyntaxProvider()
                    Exit Select
            End Select

            FillVersions()
        End Sub

        Private Sub BoxServerVersion_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            If TypeOf _connectionInfo.SyntaxProvider Is FirebirdSyntaxProvider Then
                If CStr(BoxServerVersion.SelectedItem) = "Firebird 1.0" Then
                    CType(_connectionInfo.SyntaxProvider, FirebirdSyntaxProvider).ServerVersion = FirebirdVersion.Firebird10
                ElseIf CStr(BoxServerVersion.SelectedItem) = "Firebird 1.5" Then
                    CType(_connectionInfo.SyntaxProvider, FirebirdSyntaxProvider).ServerVersion = FirebirdVersion.Firebird15
                End If

                If CStr(BoxServerVersion.SelectedItem) = "Firebird 2.0" Then
                    CType(_connectionInfo.SyntaxProvider, FirebirdSyntaxProvider).ServerVersion = FirebirdVersion.Firebird20
                End If

                If CStr(BoxServerVersion.SelectedItem) = "Firebird 2.5" Then
                    CType(_connectionInfo.SyntaxProvider, FirebirdSyntaxProvider).ServerVersion = FirebirdVersion.Firebird25
                End If
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is InformixSyntaxProvider Then
                If CStr(BoxServerVersion.SelectedItem) = "Informix 8" Then
                    CType(_connectionInfo.SyntaxProvider, InformixSyntaxProvider).ServerVersion = InformixVersion.DS8
                ElseIf CStr(BoxServerVersion.SelectedItem) = "Informix 9" Then
                    CType(_connectionInfo.SyntaxProvider, InformixSyntaxProvider).ServerVersion = InformixVersion.DS9
                End If

                If CStr(BoxServerVersion.SelectedItem) = "Informix 10" Then
                    CType(_connectionInfo.SyntaxProvider, InformixSyntaxProvider).ServerVersion = InformixVersion.DS10
                End If
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is MSAccessSyntaxProvider Then
                If CStr(BoxServerVersion.SelectedItem) = "MS Jet 3" Then
                    CType(_connectionInfo.SyntaxProvider, MSAccessSyntaxProvider).ServerVersion = MSAccessServerVersion.MSJET3
                ElseIf CStr(BoxServerVersion.SelectedItem) = "MS Jet 4" Then
                    CType(_connectionInfo.SyntaxProvider, MSAccessSyntaxProvider).ServerVersion = MSAccessServerVersion.MSJET4
                End If
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is MSSQLSyntaxProvider Then
                If CStr(BoxServerVersion.SelectedItem) = "Auto" Then
                    CType(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.Auto
                ElseIf CStr(BoxServerVersion.SelectedItem) = "SQL Server 7" Then
                    CType(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL7
                ElseIf CStr(BoxServerVersion.SelectedItem) = "SQL Server 2000" Then
                    CType(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2000
                ElseIf CStr(BoxServerVersion.SelectedItem) = "SQL Server 2005" Then
                    CType(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2005
                End If
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is MySQLSyntaxProvider Then
                If CStr(BoxServerVersion.SelectedItem) = "3.0" Then
                    CType(_connectionInfo.SyntaxProvider, MySQLSyntaxProvider).ServerVersionInt = 39999
                ElseIf CStr(BoxServerVersion.SelectedItem) = "4.0" Then
                    CType(_connectionInfo.SyntaxProvider, MySQLSyntaxProvider).ServerVersionInt = 49999
                ElseIf CStr(BoxServerVersion.SelectedItem) = "5.0" Then
                    CType(_connectionInfo.SyntaxProvider, MySQLSyntaxProvider).ServerVersionInt = 50000
                End If
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is OracleSyntaxProvider Then
                If CStr(BoxServerVersion.SelectedItem) = "Oracle 7" Then
                    CType(_connectionInfo.SyntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle7
                ElseIf CStr(BoxServerVersion.SelectedItem) = "Oracle 8" Then
                    CType(_connectionInfo.SyntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle8
                ElseIf CStr(BoxServerVersion.SelectedItem) = "Oracle 9" Then
                    CType(_connectionInfo.SyntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle9
                ElseIf CStr(BoxServerVersion.SelectedItem) = "Oracle 10" Then
                    CType(_connectionInfo.SyntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle10
                End If
            ElseIf TypeOf _connectionInfo.SyntaxProvider Is SybaseSyntaxProvider Then
                If CStr(BoxServerVersion.SelectedItem) = "ASE" Then
                    CType(_connectionInfo.SyntaxProvider, SybaseSyntaxProvider).ServerVersion = SybaseServerVersion.SybaseASE
                ElseIf CStr(BoxServerVersion.SelectedItem) = "SQL Anywhere" Then
                    CType(_connectionInfo.SyntaxProvider, SybaseSyntaxProvider).ServerVersion = SybaseServerVersion.SybaseASA
                End If
            End If
        End Sub
    End Class
End Namespace
