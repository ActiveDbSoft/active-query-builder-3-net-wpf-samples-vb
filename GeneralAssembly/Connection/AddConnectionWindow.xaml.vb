'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports ActiveQueryBuilder.Core
Imports Connection.FrameConnection
Imports System
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Threading

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
			If DialogResult.Equals(True) Then
				If _currentConnectionFrame IsNot Nothing AndAlso _currentConnectionFrame.TestConnection() Then
					_connectionInfo.Name = TextBoxConnectionName.Text
					If Not _connectionInfo.IsXmlFile Then
						'previous version of this demo uses deprecated System.Data.OracleClient
						' current version uses Oracle.ManagedDataAccess.Client which doesn't support "Integrated Security" setting
						If _connectionInfo.Type = ConnectionTypes.Oracle Then
							_currentConnectionFrame.ConnectionString = Regex.Replace(_currentConnectionFrame.ConnectionString, "Integrated Security=.*?;", "")
						End If
						_connectionInfo.ConnectionDescriptor.MetadataProvider.Connection.ConnectionString = _currentConnectionFrame.ConnectionString
					End If
					_connectionInfo.ConnectionString = _currentConnectionFrame.ConnectionString
					e.Cancel = False
					RemoveHandler Dispatcher.CurrentDispatcher.Hooks.DispatcherInactive, AddressOf Hooks_DispatcherInactive
				Else
					e.Cancel = True
				End If
			End If

			MyBase.OnClosing(e)
		End Sub

'INSTANT VB NOTE: The parameter connectionInfo was renamed since it may cause conflicts with calls to static members of the user-defined type with this name:
		Public Sub New(connectionInfo_Renamed As ConnectionInfo)
			InitializeComponent()

			Debug.Assert(connectionInfo_Renamed IsNot Nothing)

			_connectionInfo = connectionInfo_Renamed
			TextBoxConnectionName.Text = connectionInfo_Renamed.Name

			If Not String.IsNullOrEmpty(connectionInfo_Renamed.ConnectionString) Then

				If Not connectionInfo_Renamed.IsXmlFile Then
					rbMSSQL.IsEnabled = (connectionInfo_Renamed.Type = ConnectionTypes.MSSQL)
					rbMSAccess.IsEnabled = (connectionInfo_Renamed.Type = ConnectionTypes.MSAccess)
					rbOracle.IsEnabled = (connectionInfo_Renamed.Type = ConnectionTypes.Oracle)
					rbMySQL.IsEnabled = (connectionInfo_Renamed.Type = ConnectionTypes.MySQL)
					rbPostrgeSQL.IsEnabled = (connectionInfo_Renamed.Type = ConnectionTypes.PostgreSQL)
					rbOLEDB.IsEnabled = (connectionInfo_Renamed.Type = ConnectionTypes.OLEDB)
					rbODBC.IsEnabled = (connectionInfo_Renamed.Type = ConnectionTypes.ODBC)
				End If
			End If

			If connectionInfo_Renamed.IsXmlFile Then
				rbOLEDB.IsEnabled = False
				rbODBC.IsEnabled = False
			End If

			rbMSSQL.IsChecked = (connectionInfo_Renamed.Type = ConnectionTypes.MSSQL)
			rbMSAccess.IsChecked = (connectionInfo_Renamed.Type = ConnectionTypes.MSAccess)
			rbOracle.IsChecked = (connectionInfo_Renamed.Type = ConnectionTypes.Oracle)
			rbMySQL.IsChecked = (connectionInfo_Renamed.Type = ConnectionTypes.MySQL)
			rbPostrgeSQL.IsChecked = (connectionInfo_Renamed.Type = ConnectionTypes.PostgreSQL)
			rbOLEDB.IsChecked = (connectionInfo_Renamed.Type = ConnectionTypes.OLEDB)
			rbODBC.IsChecked = (connectionInfo_Renamed.Type = ConnectionTypes.ODBC)

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
				Select Case _connectionInfo.Type
					Case ConnectionTypes.MSSQL
						_connectionInfo.SyntaxProvider = New MSSQLSyntaxProvider()
					Case ConnectionTypes.MSAccess
						_connectionInfo.SyntaxProvider = New MSAccessSyntaxProvider()
					Case ConnectionTypes.Oracle
						_connectionInfo.SyntaxProvider = New OracleSyntaxProvider()
					Case ConnectionTypes.MySQL
						_connectionInfo.SyntaxProvider = New MySQLSyntaxProvider()
					Case ConnectionTypes.PostgreSQL
						_connectionInfo.SyntaxProvider = New PostgreSQLSyntaxProvider()
					Case ConnectionTypes.OLEDB
						_connectionInfo.SyntaxProvider = New SQL92SyntaxProvider()
					Case ConnectionTypes.ODBC
						_connectionInfo.SyntaxProvider = New SQL92SyntaxProvider()
				End Select
			End If

			If _connectionInfo.Type = ConnectionTypes.ODBC OrElse _connectionInfo.Type = ConnectionTypes.OLEDB Then
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

				Dim firebirdSyntaxProvider = CType(_connectionInfo.SyntaxProvider, FirebirdSyntaxProvider)

				Select Case firebirdSyntaxProvider.ServerVersion
					Case FirebirdVersion.Firebird10
						BoxServerVersion.SelectedItem = "Firebird 1.0"
					Case FirebirdVersion.Firebird15
						BoxServerVersion.SelectedItem = "Firebird 1.5"
					Case FirebirdVersion.Firebird20
						BoxServerVersion.SelectedItem = "Firebird 2.0"
					Case FirebirdVersion.Firebird25
						BoxServerVersion.SelectedItem = "Firebird 2.5"
				End Select
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is DB2SyntaxProvider Then
				BoxServerVersion.IsEnabled = False
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is InformixSyntaxProvider Then
				BoxServerVersion.IsEnabled = True
				BoxServerVersion.Items.Add("Informix 8")
				BoxServerVersion.Items.Add("Informix 9")
				BoxServerVersion.Items.Add("Informix 10")
				BoxServerVersion.Items.Add("Informix 11")

				Dim informixSyntaxProvider = CType(_connectionInfo.SyntaxProvider, InformixSyntaxProvider)

				Select Case informixSyntaxProvider.ServerVersion
					Case InformixVersion.DS8
						BoxServerVersion.SelectedItem = "Informix 8"
					Case InformixVersion.DS9
						BoxServerVersion.SelectedItem = "Informix 9"
					Case InformixVersion.DS10
						BoxServerVersion.SelectedItem = "Informix 10"
					Case Else
						BoxServerVersion.SelectedItem = "Informix 11"
				End Select
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is MSAccessSyntaxProvider Then
				BoxServerVersion.IsEnabled = True
				BoxServerVersion.Items.Add("Access 97")
				BoxServerVersion.Items.Add("Access 2000 and newer")

				Dim accessSyntaxProvider = CType(_connectionInfo.SyntaxProvider, MSAccessSyntaxProvider)

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

				Dim mssqlSyntaxProvider = CType(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider)

				Select Case mssqlSyntaxProvider.ServerVersion
					Case MSSQLServerVersion.MSSQL7
						BoxServerVersion.SelectedItem = "SQL Server 7"
					Case MSSQLServerVersion.MSSQL2000
						BoxServerVersion.SelectedItem = "SQL Server 2000"
					Case MSSQLServerVersion.MSSQL2005
						BoxServerVersion.SelectedItem = "SQL Server 2005"
					Case MSSQLServerVersion.MSSQL2012
						BoxServerVersion.SelectedItem = "SQL Server 2012"
					Case MSSQLServerVersion.MSSQL2014
						BoxServerVersion.SelectedItem = "SQL Server 2014"
					Case MSSQLServerVersion.MSSQL2016
						BoxServerVersion.SelectedItem = "SQL Server 2016"
					Case MSSQLServerVersion.MSSQL2017
						BoxServerVersion.SelectedItem = "SQL Server 2017"
					Case MSSQLServerVersion.MSSQL2019
						BoxServerVersion.SelectedItem = "SQL Server 2019"
					Case Else
						BoxServerVersion.SelectedItem = "Auto"
				End Select
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is MySQLSyntaxProvider Then
				BoxServerVersion.IsEnabled = True
				BoxServerVersion.Items.Add("3.0")
				BoxServerVersion.Items.Add("4.0")
				BoxServerVersion.Items.Add("5.0")
				BoxServerVersion.Items.Add("8.0+")

				Dim mySqlSyntaxProvider = CType(_connectionInfo.SyntaxProvider, MySQLSyntaxProvider)

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

				Dim oracleSyntaxProvider = CType(_connectionInfo.SyntaxProvider, OracleSyntaxProvider)

				Select Case oracleSyntaxProvider.ServerVersion
					Case OracleServerVersion.Oracle7
						BoxServerVersion.SelectedItem = "Oracle 7"
					Case OracleServerVersion.Oracle8
						BoxServerVersion.SelectedItem = "Oracle 8"
					Case OracleServerVersion.Oracle9
						BoxServerVersion.SelectedItem = "Oracle 9"
					Case OracleServerVersion.Oracle10
						BoxServerVersion.SelectedItem = "Oracle 10"
					Case OracleServerVersion.Oracle11
						BoxServerVersion.SelectedItem = "Oracle 11g"
					Case OracleServerVersion.Oracle12
						BoxServerVersion.SelectedItem = "Oracle 12c"
					Case OracleServerVersion.Oracle18
						BoxServerVersion.SelectedItem = "Oracle 18c"
					Case OracleServerVersion.Oracle19
						BoxServerVersion.SelectedItem = "Oracle 19c"
					Case Else
						BoxServerVersion.SelectedItem = "Oracle 18c"
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

				Dim sybaseSyntaxProvider = CType(_connectionInfo.SyntaxProvider, SybaseSyntaxProvider)

				Select Case sybaseSyntaxProvider.ServerVersion
					Case SybaseServerVersion.SybaseASE
						BoxServerVersion.SelectedItem = "ASE"
					Case SybaseServerVersion.SybaseASA
						BoxServerVersion.SelectedItem = "SQL Anywhere"
					Case SybaseServerVersion.SybaseIQ
						BoxServerVersion.SelectedItem = "SAP IQ"
				End Select
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is VistaDBSyntaxProvider Then
				BoxServerVersion.IsEnabled = False
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is GenericSyntaxProvider Then
				BoxServerVersion.IsEnabled = False
			End If
		End Sub

		Private Sub ConnectionTypeChanged(sender As Object, e As RoutedEventArgs)
			If Not DirectCast(sender, RadioButton).IsChecked.Equals(True) Then
				Return
			End If

			Dim connectionType = ConnectionTypes.MSSQL

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

			If connectionType <> _connectionInfo.Type Then
				_connectionInfo.Type = connectionType

				If Not _connectionInfo.IsXmlFile Then
					SetActiveConnectionTypeFrame()
				End If
			End If

			Select Case _connectionInfo.Type
				Case ConnectionTypes.MSSQL
					_connectionInfo.SyntaxProvider = New MSSQLSyntaxProvider()
				Case ConnectionTypes.MSAccess
					_connectionInfo.SyntaxProvider = New MSAccessSyntaxProvider()
				Case ConnectionTypes.Oracle
					_connectionInfo.SyntaxProvider = New OracleSyntaxProvider()
				Case ConnectionTypes.MySQL
					_connectionInfo.SyntaxProvider = New MySQLSyntaxProvider()
				Case ConnectionTypes.PostgreSQL
					_connectionInfo.SyntaxProvider = New PostgreSQLSyntaxProvider()
				Case ConnectionTypes.OLEDB
					_connectionInfo.SyntaxProvider = New SQL92SyntaxProvider()
				Case ConnectionTypes.ODBC
					_connectionInfo.SyntaxProvider = New SQL92SyntaxProvider()
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
				Select Case _connectionInfo.Type
					Case ConnectionTypes.MSSQL
						_currentConnectionFrame = New MSSQLConnectionFrame(_connectionInfo.ConnectionString)
					Case ConnectionTypes.MSAccess
						_currentConnectionFrame = New MSAccessConnectionFrame(_connectionInfo.ConnectionString)
					Case ConnectionTypes.Oracle
						_currentConnectionFrame = New OracleConnectionFrame(_connectionInfo.ConnectionString)
					Case ConnectionTypes.MySQL
						_currentConnectionFrame = New MySQLConnectionFrame(_connectionInfo.ConnectionString)
					Case ConnectionTypes.PostgreSQL
						_currentConnectionFrame = New PostgreSQLConnectionFrame(_connectionInfo.ConnectionString)
					Case ConnectionTypes.OLEDB
						_currentConnectionFrame = New OLEDBConnectionFrame(_connectionInfo.ConnectionString)
					Case ConnectionTypes.ODBC
						_currentConnectionFrame = New ODBCConnectionFrame(_connectionInfo.ConnectionString)
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
			Dim syntaxProvider = TryCast(Activator.CreateInstance(syntaxType), BaseSyntaxProvider)
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
			Select Case (CStr(BoxSyntaxProvider.SelectedItem))
				Case "ANSI SQL-2003"
					_connectionInfo.SyntaxProvider = New SQL2003SyntaxProvider()
				Case "ANSI SQL-92"
					_connectionInfo.SyntaxProvider = New SQL92SyntaxProvider()
				Case "ANSI SQL-89"
					_connectionInfo.SyntaxProvider = New SQL89SyntaxProvider()
				Case "Firebird"
					_connectionInfo.SyntaxProvider = New FirebirdSyntaxProvider()
				Case "IBM DB2"
					_connectionInfo.SyntaxProvider = New DB2SyntaxProvider()
				Case "IBM Informix"
					_connectionInfo.SyntaxProvider = New InformixSyntaxProvider()
				Case "Microsoft Access"
					_connectionInfo.SyntaxProvider = New MSAccessSyntaxProvider()
				Case "Microsoft SQL Server"
					_connectionInfo.SyntaxProvider = New MSSQLSyntaxProvider()
				Case "MySQL"
					_connectionInfo.SyntaxProvider = New MySQLSyntaxProvider()
				Case "Oracle"
					_connectionInfo.SyntaxProvider = New OracleSyntaxProvider()
				Case "PostgreSQL"
					_connectionInfo.SyntaxProvider = New PostgreSQLSyntaxProvider()
				Case "SQLite"
					_connectionInfo.SyntaxProvider = New SQLiteSyntaxProvider()
				Case "Sybase"
					_connectionInfo.SyntaxProvider = New SybaseSyntaxProvider()
				Case "VistaDB"
					_connectionInfo.SyntaxProvider = New VistaDBSyntaxProvider()
				Case "Universal"
					_connectionInfo.SyntaxProvider = New GenericSyntaxProvider()
			End Select

			FillVersions()
		End Sub

		Private Sub BoxServerVersion_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
			If BoxServerVersion.SelectedItem Is Nothing Then
				Return
			End If

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
				If CStr(BoxServerVersion.SelectedItem) = "Access 97" Then
					CType(_connectionInfo.SyntaxProvider, MSAccessSyntaxProvider).ServerVersion = MSAccessServerVersion.MSJET3
				ElseIf CStr(BoxServerVersion.SelectedItem) = "Access 2000 and newer" Then
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
				ElseIf CStr(BoxServerVersion.SelectedItem) = "SQL Server 2008" Then
					CType(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2008
				ElseIf CStr(BoxServerVersion.SelectedItem) = "SQL Server 2012" Then
					CType(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2012
				ElseIf CStr(BoxServerVersion.SelectedItem) = "SQL Server 2014" Then
					CType(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2014
				ElseIf CStr(BoxServerVersion.SelectedItem) = "SQL Server 2016" Then
					CType(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2016
				ElseIf CStr(BoxServerVersion.SelectedItem) = "SQL Server 2017" Then
					CType(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2017
				ElseIf CStr(BoxServerVersion.SelectedItem) = "SQL Server 2019" Then
					CType(_connectionInfo.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2017
				End If
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is MySQLSyntaxProvider Then
				If CStr(BoxServerVersion.SelectedItem) = "3.0" Then
					CType(_connectionInfo.SyntaxProvider, MySQLSyntaxProvider).ServerVersionInt = 39999
				ElseIf CStr(BoxServerVersion.SelectedItem) = "4.0" Then
					CType(_connectionInfo.SyntaxProvider, MySQLSyntaxProvider).ServerVersionInt = 49999
				ElseIf CStr(BoxServerVersion.SelectedItem) = "5.0" Then
					CType(_connectionInfo.SyntaxProvider, MySQLSyntaxProvider).ServerVersionInt = 50000
				ElseIf CStr(BoxServerVersion.SelectedItem) = "8.0+" Then
					CType(_connectionInfo.SyntaxProvider, MySQLSyntaxProvider).ServerVersionInt = 80012
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
				ElseIf CStr(BoxServerVersion.SelectedItem) = "Oracle 11g" Then
					CType(_connectionInfo.SyntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle11
				ElseIf CStr(BoxServerVersion.SelectedItem) = "Oracle 12c" Then
					CType(_connectionInfo.SyntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle12
				ElseIf CStr(BoxServerVersion.SelectedItem) = "Oracle 18c" Then
					CType(_connectionInfo.SyntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle18
				ElseIf CStr(BoxServerVersion.SelectedItem) = "Oracle 19c" Then
					CType(_connectionInfo.SyntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle19
				End If
			ElseIf TypeOf _connectionInfo.SyntaxProvider Is SybaseSyntaxProvider Then
				If CStr(BoxServerVersion.SelectedItem) = "ASE" Then
					CType(_connectionInfo.SyntaxProvider, SybaseSyntaxProvider).ServerVersion = SybaseServerVersion.SybaseASE
				ElseIf CStr(BoxServerVersion.SelectedItem) = "SQL Anywhere" Then
					CType(_connectionInfo.SyntaxProvider, SybaseSyntaxProvider).ServerVersion = SybaseServerVersion.SybaseASA
				End If
			End If

			_currentConnectionFrame.SetServerType(TryCast(BoxServerVersion.SelectedItem, String))
		End Sub
	End Class
End Namespace
