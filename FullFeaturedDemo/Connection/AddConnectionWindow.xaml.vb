'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
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
	Public Partial Class AddConnectionWindow
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

		Private Function SyntaxToString(syntax As BaseSyntaxProvider) As String
			If TypeOf syntax Is SQL2003SyntaxProvider Then
				Return "ANSI SQL-2003"
			ElseIf TypeOf syntax Is SQL92SyntaxProvider Then
				Return "ANSI SQL-92"
			ElseIf TypeOf syntax Is SQL89SyntaxProvider Then
				Return "ANSI SQL-89"
			ElseIf TypeOf syntax Is FirebirdSyntaxProvider Then
				Return "Firebird"
			ElseIf TypeOf syntax Is DB2SyntaxProvider Then
				Return "IBM DB2"
			ElseIf TypeOf syntax Is InformixSyntaxProvider Then
				Return "IBM Informix"
			ElseIf TypeOf syntax Is MSAccessSyntaxProvider Then
				Return "Microsoft Access"
			ElseIf TypeOf syntax Is MSSQLSyntaxProvider Then
				Return "Microsoft SQL Server"
			ElseIf TypeOf syntax Is MySQLSyntaxProvider Then
				Return "MySQL"
			ElseIf TypeOf syntax Is OracleSyntaxProvider Then
				Return "Oracle"
			ElseIf TypeOf syntax Is PostgreSQLSyntaxProvider Then
				Return "PostgreSQL"
			ElseIf TypeOf syntax Is SQLiteSyntaxProvider Then
				Return "SQLite"
			ElseIf TypeOf syntax Is SybaseSyntaxProvider Then
				Return "Sybase"
			ElseIf TypeOf syntax Is VistaDBSyntaxProvider Then
				Return "VistaDB"
			ElseIf TypeOf syntax Is GenericSyntaxProvider Then
				Return "Universal"
			End If

			Return String.Empty
		End Function

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

			If _connectionInfo.ConnectionType = ConnectionTypes.ODBC OrElse _connectionInfo.ConnectionType = ConnectionTypes.OLEDB Then
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
				BoxSyntaxProvider.SelectedItem = SyntaxToString(_connectionInfo.SyntaxProvider)
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is SQL2003SyntaxProvider Then
				BoxSyntaxProvider.Items.Add(SyntaxToString(_connectionInfo.SyntaxProvider))
				BoxSyntaxProvider.SelectedItem = SyntaxToString(_connectionInfo.SyntaxProvider)

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
				BoxSyntaxProvider.Items.Add(SyntaxToString(_connectionInfo.SyntaxProvider))
				BoxSyntaxProvider.SelectedItem = SyntaxToString(_connectionInfo.SyntaxProvider)

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
				BoxSyntaxProvider.Items.Add(SyntaxToString(_connectionInfo.SyntaxProvider))
				BoxSyntaxProvider.SelectedItem = SyntaxToString(_connectionInfo.SyntaxProvider)

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
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is GenericSyntaxProvider Then
				BoxSyntaxProvider.Items.Add(SyntaxToString(_connectionInfo.SyntaxProvider))
				BoxSyntaxProvider.SelectedItem = SyntaxToString(_connectionInfo.SyntaxProvider)

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
			Else
				BoxSyntaxProvider.Items.Add(SyntaxToString(_connectionInfo.SyntaxProvider))
				BoxSyntaxProvider.SelectedItem = SyntaxToString(_connectionInfo.SyntaxProvider)
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

                Dim firebirdSyntaxProvider As FirebirdSyntaxProvider = DirectCast(_connectionInfo.SyntaxProvider, FirebirdSyntaxProvider)

                Select Case firebirdSyntaxProvider.ServerVersion
					Case FirebirdVersion.Firebird10
						BoxServerVersion.SelectedItem = "Firebird 1.0"
						Exit Select
					Case FirebirdVersion.Firebird15
						BoxServerVersion.SelectedItem = "Firebird 1.5"
						Exit Select
					Case FirebirdVersion.Firebird20
						BoxServerVersion.SelectedItem = "Firebird 2.0"
						Exit Select
					Case FirebirdVersion.Firebird25
						BoxServerVersion.SelectedItem = "Firebird 2.5"
						Exit Select
				End Select
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is DB2SyntaxProvider Then
				BoxServerVersion.IsEnabled = False
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is InformixSyntaxProvider Then
				BoxServerVersion.IsEnabled = True
				BoxServerVersion.Items.Add("Informix 8")
				BoxServerVersion.Items.Add("Informix 9")
				BoxServerVersion.Items.Add("Informix 10")
				BoxServerVersion.Items.Add("Informix 11")

                Dim informixSyntaxProvider As InformixSyntaxProvider = DirectCast(_connectionInfo.SyntaxProvider, InformixSyntaxProvider)

                Select Case informixSyntaxProvider.ServerVersion
					Case InformixVersion.DS8
						BoxServerVersion.SelectedItem = "Informix 8"
						Exit Select
					Case InformixVersion.DS9
						BoxServerVersion.SelectedItem = "Informix 9"
						Exit Select
					Case InformixVersion.DS10
						BoxServerVersion.SelectedItem = "Informix 10"
						Exit Select
					Case Else
						BoxServerVersion.SelectedItem = "Informix 11"
						Exit Select
				End Select
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is MSAccessSyntaxProvider Then
				BoxServerVersion.IsEnabled = True
				BoxServerVersion.Items.Add("Access 97")
				BoxServerVersion.Items.Add("Access 2000 and newer")

                Dim accessSyntaxProvider As MSAccessSyntaxProvider = DirectCast(_connectionInfo.SyntaxProvider, MSAccessSyntaxProvider)

                BoxServerVersion.SelectedItem = If(accessSyntaxProvider.ServerVersion = MSAccessServerVersion.MSJET3, "Access 97", "Access 2000 and newer")
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is MSSQLSyntaxProvider Then
				BoxServerVersion.IsEnabled = True
				BoxServerVersion.Items.Add("Auto")
				BoxServerVersion.Items.Add("SQL Server 7")
				BoxServerVersion.Items.Add("SQL Server 2000")
				BoxServerVersion.Items.Add("SQL Server 2005")
				BoxServerVersion.Items.Add("SQL Server 2012")
				BoxServerVersion.Items.Add("SQL Server 2014")
				BoxServerVersion.Items.Add("SQL Server 2016")
				BoxServerVersion.Items.Add("SQL Server 2017")
				BoxServerVersion.Items.Add("SQL Server 2019")

                Dim mssqlSyntaxProvider As MSSQLSyntaxProvider = DirectCast(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider)

                Select Case mssqlSyntaxProvider.ServerVersion
					Case MSSQLServerVersion.MSSQL7
						BoxServerVersion.SelectedItem = "SQL Server 7"
						Exit Select
					Case MSSQLServerVersion.MSSQL2000
						BoxServerVersion.SelectedItem = "SQL Server 2000"
						Exit Select
					Case MSSQLServerVersion.MSSQL2005
						BoxServerVersion.SelectedItem = "SQL Server 2005"
						Exit Select
					Case MSSQLServerVersion.MSSQL2012
						BoxServerVersion.SelectedItem = "SQL Server 2012"
						Exit Select
					Case MSSQLServerVersion.MSSQL2014
						BoxServerVersion.SelectedItem = "SQL Server 2014"
						Exit Select
					Case MSSQLServerVersion.MSSQL2016
						BoxServerVersion.SelectedItem = "SQL Server 2016"
						Exit Select
					Case MSSQLServerVersion.MSSQL2017
						BoxServerVersion.SelectedItem = "SQL Server 2017"
						Exit Select
					Case MSSQLServerVersion.MSSQL2019
						BoxServerVersion.SelectedItem = "SQL Server 2019"
						Exit Select
					Case Else
						BoxServerVersion.SelectedItem = "Auto"
						Exit Select
				End Select
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is MySQLSyntaxProvider Then
				BoxServerVersion.IsEnabled = True
				BoxServerVersion.Items.Add("3.0")
				BoxServerVersion.Items.Add("4.0")
				BoxServerVersion.Items.Add("5.0")
				BoxServerVersion.Items.Add("8.0+")

                Dim mySqlSyntaxProvider As MySQLSyntaxProvider = DirectCast(_connectionInfo.SyntaxProvider, MySQLSyntaxProvider)

                If mySqlSyntaxProvider.ServerVersionInt < 40000 Then
					BoxServerVersion.SelectedItem = "3.0"
				ElseIf mySqlSyntaxProvider.ServerVersionInt < 50000 Then
					BoxServerVersion.SelectedItem = "4.0"
				ElseIf mySqlSyntaxProvider.ServerVersionInt < 80000 Then
					BoxServerVersion.SelectedItem = "5.0"
				Else
					BoxServerVersion.SelectedItem = "8.0+"
				End If
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is OracleSyntaxProvider Then
				BoxServerVersion.IsEnabled = True
				BoxServerVersion.Items.Add("Oracle 7")
				BoxServerVersion.Items.Add("Oracle 8")
				BoxServerVersion.Items.Add("Oracle 9")
				BoxServerVersion.Items.Add("Oracle 10")
				BoxServerVersion.Items.Add("Oracle 11g")
				BoxServerVersion.Items.Add("Oracle 12c")
				BoxServerVersion.Items.Add("Oracle 18c")
				BoxServerVersion.Items.Add("Oracle 19c")

                Dim oracleSyntaxProvider As OracleSyntaxProvider = DirectCast(_connectionInfo.SyntaxProvider, OracleSyntaxProvider)

                Select Case oracleSyntaxProvider.ServerVersion
					Case OracleServerVersion.Oracle7
						BoxServerVersion.SelectedItem = "Oracle 7"
						Exit Select
					Case OracleServerVersion.Oracle8
						BoxServerVersion.SelectedItem = "Oracle 8"
						Exit Select
					Case OracleServerVersion.Oracle9
						BoxServerVersion.SelectedItem = "Oracle 9"
						Exit Select
					Case OracleServerVersion.Oracle10
						BoxServerVersion.SelectedItem = "Oracle 10"
						Exit Select
					Case OracleServerVersion.Oracle11
						BoxServerVersion.SelectedItem = "Oracle 11g"
						Exit Select
					Case OracleServerVersion.Oracle12
						BoxServerVersion.SelectedItem = "Oracle 12c"
						Exit Select
					Case OracleServerVersion.Oracle18
						BoxServerVersion.SelectedItem = "Oracle 18c"
						Exit Select
					Case OracleServerVersion.Oracle19
						BoxServerVersion.SelectedItem = "Oracle 19c"
						Exit Select
					Case Else
						BoxServerVersion.SelectedItem = "Oracle 18c"
						Exit Select
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

                Dim sybaseSyntaxProvider As SybaseSyntaxProvider = DirectCast(_connectionInfo.SyntaxProvider, SybaseSyntaxProvider)

                Select Case sybaseSyntaxProvider.ServerVersion
					Case SybaseServerVersion.SybaseASE
						BoxServerVersion.SelectedItem = "ASE"
						Exit Select
					Case SybaseServerVersion.SybaseASA
						BoxServerVersion.SelectedItem = "SQL Anywhere"
						Exit Select
					Case SybaseServerVersion.SybaseIQ
						BoxServerVersion.SelectedItem = "SAP IQ"
						Exit Select
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
				RemoveHandler _currentConnectionFrame.OnSyntaxProviderDetected, AddressOf CurrentConnectionFrame_SyntaxProviderDetected
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
				AddHandler _currentConnectionFrame.OnSyntaxProviderDetected, AddressOf CurrentConnectionFrame_SyntaxProviderDetected
			End If
		End Sub

		Private Sub CurrentConnectionFrame_SyntaxProviderDetected(syntaxType As Type)
            Dim syntaxProvider As BaseSyntaxProvider = TryCast(Activator.CreateInstance(syntaxType), BaseSyntaxProvider)
            BoxSyntaxProvider.SelectedItem = SyntaxToString(syntaxProvider)
			FillVersions()
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
			If BoxServerVersion.SelectedItem Is Nothing Then
				Return
			End If

			If TypeOf _connectionInfo.SyntaxProvider Is FirebirdSyntaxProvider Then
				If DirectCast(BoxServerVersion.SelectedItem, String) = "Firebird 1.0" Then
					DirectCast(_connectionInfo.SyntaxProvider, FirebirdSyntaxProvider).ServerVersion = FirebirdVersion.Firebird10
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "Firebird 1.5" Then
					DirectCast(_connectionInfo.SyntaxProvider, FirebirdSyntaxProvider).ServerVersion = FirebirdVersion.Firebird15
				End If
				If DirectCast(BoxServerVersion.SelectedItem, String) = "Firebird 2.0" Then
					DirectCast(_connectionInfo.SyntaxProvider, FirebirdSyntaxProvider).ServerVersion = FirebirdVersion.Firebird20
				End If
				If DirectCast(BoxServerVersion.SelectedItem, String) = "Firebird 2.5" Then
					DirectCast(_connectionInfo.SyntaxProvider, FirebirdSyntaxProvider).ServerVersion = FirebirdVersion.Firebird25
				End If
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is InformixSyntaxProvider Then
				If DirectCast(BoxServerVersion.SelectedItem, String) = "Informix 8" Then
					DirectCast(_connectionInfo.SyntaxProvider, InformixSyntaxProvider).ServerVersion = InformixVersion.DS8
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "Informix 9" Then
					DirectCast(_connectionInfo.SyntaxProvider, InformixSyntaxProvider).ServerVersion = InformixVersion.DS9
				End If
				If DirectCast(BoxServerVersion.SelectedItem, String) = "Informix 10" Then
					DirectCast(_connectionInfo.SyntaxProvider, InformixSyntaxProvider).ServerVersion = InformixVersion.DS10
				End If
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is MSAccessSyntaxProvider Then
				If DirectCast(BoxServerVersion.SelectedItem, String) = "Access 97" Then
					DirectCast(_connectionInfo.SyntaxProvider, MSAccessSyntaxProvider).ServerVersion = MSAccessServerVersion.MSJET3
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "Access 2000 and newer" Then
					DirectCast(_connectionInfo.SyntaxProvider, MSAccessSyntaxProvider).ServerVersion = MSAccessServerVersion.MSJET4
				End If
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is MSSQLSyntaxProvider Then
				If DirectCast(BoxServerVersion.SelectedItem, String) = "Auto" Then
					DirectCast(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.Auto
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "SQL Server 7" Then
					DirectCast(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL7
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "SQL Server 2000" Then
					DirectCast(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2000
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "SQL Server 2005" Then
					DirectCast(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2005
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "SQL Server 2008" Then
					DirectCast(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2008
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "SQL Server 2012" Then
					DirectCast(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2012
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "SQL Server 2014" Then
					DirectCast(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2014
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "SQL Server 2016" Then
					DirectCast(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2016
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "SQL Server 2017" Then
					DirectCast(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2017
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "SQL Server 2019" Then
					DirectCast(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2019
				End If
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is MySQLSyntaxProvider Then
				If DirectCast(BoxServerVersion.SelectedItem, String) = "3.0" Then
					DirectCast(_connectionInfo.SyntaxProvider, MySQLSyntaxProvider).ServerVersionInt = 39999
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "4.0" Then
					DirectCast(_connectionInfo.SyntaxProvider, MySQLSyntaxProvider).ServerVersionInt = 49999
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "5.0" Then
					DirectCast(_connectionInfo.SyntaxProvider, MySQLSyntaxProvider).ServerVersionInt = 50000
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "8.0+" Then
					DirectCast(_connectionInfo.SyntaxProvider, MySQLSyntaxProvider).ServerVersionInt = 80012
				End If
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is OracleSyntaxProvider Then
				If DirectCast(BoxServerVersion.SelectedItem, String) = "Oracle 7" Then
					DirectCast(_connectionInfo.SyntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle7
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "Oracle 8" Then
					DirectCast(_connectionInfo.SyntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle8
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "Oracle 9" Then
					DirectCast(_connectionInfo.SyntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle9
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "Oracle 10" Then
					DirectCast(_connectionInfo.SyntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle10
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "Oracle 11g" Then
					DirectCast(_connectionInfo.SyntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle11
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "Oracle 12c" Then
					DirectCast(_connectionInfo.SyntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle12
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "Oracle 18c" Then
					DirectCast(_connectionInfo.SyntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle18
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "Oracle 19c" Then
					DirectCast(_connectionInfo.SyntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle19
				End If
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is SybaseSyntaxProvider Then
				If DirectCast(BoxServerVersion.SelectedItem, String) = "ASE" Then
					DirectCast(_connectionInfo.SyntaxProvider, SybaseSyntaxProvider).ServerVersion = SybaseServerVersion.SybaseASE
				ElseIf DirectCast(BoxServerVersion.SelectedItem, String) = "SQL Anywhere" Then
					DirectCast(_connectionInfo.SyntaxProvider, SybaseSyntaxProvider).ServerVersion = SybaseServerVersion.SybaseASA
				End If
			End If

			_currentConnectionFrame.SetServerType(TryCast(BoxServerVersion.SelectedItem, String))
		End Sub
	End Class
End Namespace
