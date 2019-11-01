﻿'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Windows.Controls
Imports ActiveQueryBuilder.Core

Namespace PropertiesForm
	''' <summary>
	''' Interaction logic for SqlSyntaxPage.xaml
	''' </summary>
	Public Partial Class SqlSyntaxPage
		Public Property Modified() As Boolean
			Get
				Return m_Modified
			End Get
			Set
				m_Modified = Value
			End Set
		End Property
		Private m_Modified As Boolean

		Private _syntaxProvider As BaseSyntaxProvider = Nothing
		Private ReadOnly _sqlContext As SQLContext
		Public Sub New()
			Modified = False
			InitializeComponent()
		End Sub

		Public Sub New(sqlQuery As SQLContext, syntaxProvider As BaseSyntaxProvider)
			Modified = False
			_syntaxProvider = syntaxProvider
			_sqlContext = sqlQuery
			InitializeComponent()

			comboIdentCaseSens.Items.Add("All identifiers are case insensitive")
			comboIdentCaseSens.Items.Add("Quoted are sensitive, Unquoted are converted to uppercase")
			comboIdentCaseSens.Items.Add("Quoted are sensitive, Unquoted are converted to lowercase")

			comboSqlDialect.Items.Add("Auto")
			comboSqlDialect.Items.Add("ANSI SQL-2003")
			comboSqlDialect.Items.Add("ANSI SQL-89")
			comboSqlDialect.Items.Add("ANSI SQL-92")
			comboSqlDialect.Items.Add("Firebird 1.0")
			comboSqlDialect.Items.Add("Firebird 1.5")
			comboSqlDialect.Items.Add("Firebird 2.0")
			comboSqlDialect.Items.Add("Firebird 2.5")
			comboSqlDialect.Items.Add("IBM DB2")
			comboSqlDialect.Items.Add("IBM Informix 8")
			comboSqlDialect.Items.Add("IBM Informix 9")
			comboSqlDialect.Items.Add("IBM Informix 10")
			comboSqlDialect.Items.Add("MS Access 97 (MS Jet 3.0)")
			comboSqlDialect.Items.Add("MS Access 2000 (MS Jet 4.0)")
			comboSqlDialect.Items.Add("MS Access XP (MS Jet 4.0)")
			comboSqlDialect.Items.Add("MS Access 2003 (MS Jet 4.0)")
			comboSqlDialect.Items.Add("MS SQL Server 7")
			comboSqlDialect.Items.Add("MS SQL Server 2000")
			comboSqlDialect.Items.Add("MS SQL Server 2005")
			comboSqlDialect.Items.Add("MS SQL Server 2008")
			comboSqlDialect.Items.Add("MS SQL Server 2012")
			comboSqlDialect.Items.Add("MS SQL Server 2014")
			comboSqlDialect.Items.Add("MS SQL Server 2016")
			comboSqlDialect.Items.Add("MS SQL Server 2017")
			comboSqlDialect.Items.Add("MS SQL Server 2019")
			comboSqlDialect.Items.Add("MS SQL Server Compact Edition")
			comboSqlDialect.Items.Add("MySQL 3.xx")
			comboSqlDialect.Items.Add("MySQL 4.0")
			comboSqlDialect.Items.Add("MySQL 4.1")
			comboSqlDialect.Items.Add("MySQL 5.0")
			comboSqlDialect.Items.Add("MySQL 8.0")
			comboSqlDialect.Items.Add("Oracle 7")
			comboSqlDialect.Items.Add("Oracle 8")
			comboSqlDialect.Items.Add("Oracle 9")
			comboSqlDialect.Items.Add("Oracle 10")
			comboSqlDialect.Items.Add("Oracle 11g")
			comboSqlDialect.Items.Add("Oracle 12c")
			comboSqlDialect.Items.Add("Oracle 18c")
			comboSqlDialect.Items.Add("Oracle 19c")
			comboSqlDialect.Items.Add("PostgreSQL")
			comboSqlDialect.Items.Add("SQLite")
			comboSqlDialect.Items.Add("Sybase ASE")
			comboSqlDialect.Items.Add("Sybase SQL Anywhere")
			comboSqlDialect.Items.Add("Teradata")
			comboSqlDialect.Items.Add("VistaDB")
			comboSqlDialect.Items.Add("Generic")

			If TypeOf _sqlContext.SyntaxProvider Is SQL92SyntaxProvider Then
				comboSqlDialect.SelectedItem = "ANSI SQL-92"
			ElseIf TypeOf _sqlContext.SyntaxProvider Is AutoSyntaxProvider Then
				comboSqlDialect.SelectedItem = "Auto"
			ElseIf TypeOf _sqlContext.SyntaxProvider Is SQL89SyntaxProvider Then
				comboSqlDialect.SelectedItem = "ANSI SQL-89"
			ElseIf TypeOf _sqlContext.SyntaxProvider Is SQL2003SyntaxProvider Then
				comboSqlDialect.SelectedItem = "ANSI SQL-2003"
			ElseIf TypeOf _sqlContext.SyntaxProvider Is FirebirdSyntaxProvider Then
				Select Case TryCast(_sqlContext.SyntaxProvider, FirebirdSyntaxProvider).ServerVersion
					Case FirebirdVersion.Firebird10
						comboSqlDialect.SelectedItem = "Firebird 1.0"
						Exit Select
					Case FirebirdVersion.Firebird15
						comboSqlDialect.SelectedItem = "Firebird 1.5"
						Exit Select
					Case FirebirdVersion.Firebird25
						comboSqlDialect.SelectedItem = "Firebird 2.5"
						Exit Select
					Case Else
						comboSqlDialect.SelectedItem = "Firebird 2.0"
						Exit Select
				End Select
			ElseIf TypeOf _sqlContext.SyntaxProvider Is DB2SyntaxProvider Then
				comboSqlDialect.SelectedItem = "IBM DB2"
			ElseIf TypeOf _sqlContext.SyntaxProvider Is InformixSyntaxProvider Then
				Select Case TryCast(_sqlContext.SyntaxProvider, InformixSyntaxProvider).ServerVersion
					Case InformixVersion.DS8
						comboSqlDialect.SelectedItem = "IBM Informix 8"
						Exit Select
					Case InformixVersion.DS9
						comboSqlDialect.SelectedItem = "IBM Informix 9"
						Exit Select
					Case Else
						comboSqlDialect.SelectedItem = "IBM Informix 10"
						Exit Select
				End Select
			ElseIf TypeOf _sqlContext.SyntaxProvider Is MSAccessSyntaxProvider Then
				Select Case TryCast(_sqlContext.SyntaxProvider, MSAccessSyntaxProvider).ServerVersion
					Case MSAccessServerVersion.MSJET3
						comboSqlDialect.SelectedItem = "MS Access 97 (MS Jet 3.0)"
						Exit Select
					Case MSAccessServerVersion.MSJET4
						comboSqlDialect.SelectedItem = "MS Access 2003 (MS Jet 4.0)"
						Exit Select
					Case Else
						comboSqlDialect.SelectedItem = "MS Access 2003 (MS Jet 4.0)"
						Exit Select
				End Select
			ElseIf TypeOf _sqlContext.SyntaxProvider Is MSSQLCESyntaxProvider Then
				comboSqlDialect.SelectedItem = "MS SQL Server Compact Edition"
			ElseIf TypeOf _sqlContext.SyntaxProvider Is MSSQLSyntaxProvider Then
				Select Case TryCast(_sqlContext.SyntaxProvider, MSSQLSyntaxProvider).ServerVersion
					Case MSSQLServerVersion.MSSQL7
						comboSqlDialect.SelectedItem = "MS SQL Server 7"
						Exit Select
					Case MSSQLServerVersion.MSSQL2000
						comboSqlDialect.SelectedItem = "MS SQL Server 2000"
						Exit Select
					Case MSSQLServerVersion.MSSQL2005
						comboSqlDialect.SelectedItem = "MS SQL Server 2005"
						Exit Select
					Case MSSQLServerVersion.MSSQL2008
						comboSqlDialect.SelectedItem = "MS SQL Server 2008"
						Exit Select
					Case MSSQLServerVersion.MSSQL2012
						comboSqlDialect.SelectedItem = "MS SQL Server 2012"
						Exit Select
					Case MSSQLServerVersion.MSSQL2014
						comboSqlDialect.SelectedItem = "MS SQL Server 2014"
						Exit Select
					Case MSSQLServerVersion.MSSQL2016
						comboSqlDialect.SelectedItem = "MS SQL Server 2016"
						Exit Select
					Case MSSQLServerVersion.MSSQL2017
						comboSqlDialect.SelectedItem = "MS SQL Server 2017"
						Exit Select
					Case MSSQLServerVersion.MSSQL2019
						comboSqlDialect.SelectedItem = "MS SQL Server 2019"
						Exit Select
					Case Else
						comboSqlDialect.SelectedItem = "MS SQL Server 2017"
						Exit Select
				End Select
			ElseIf TypeOf _sqlContext.SyntaxProvider Is MySQLSyntaxProvider Then
				If TryCast(_sqlContext.SyntaxProvider, MySQLSyntaxProvider).ServerVersionInt < 40000 Then
					comboSqlDialect.SelectedItem = "MySQL 3.xx"
				ElseIf TryCast(_sqlContext.SyntaxProvider, MySQLSyntaxProvider).ServerVersionInt <= 40099 Then
					comboSqlDialect.SelectedItem = "MySQL 4.0"
				ElseIf TryCast(_sqlContext.SyntaxProvider, MySQLSyntaxProvider).ServerVersionInt < 50000 Then
					comboSqlDialect.SelectedItem = "MySQL 4.1"
				ElseIf TryCast(_sqlContext.SyntaxProvider, MySQLSyntaxProvider).ServerVersionInt < 80000 Then
					comboSqlDialect.SelectedItem = "MySQL 5.0"
				Else
					comboSqlDialect.SelectedItem = "MySQL 8.0"
				End If
			ElseIf TypeOf _sqlContext.SyntaxProvider Is OracleSyntaxProvider Then
				Select Case TryCast(_sqlContext.SyntaxProvider, OracleSyntaxProvider).ServerVersion
					Case OracleServerVersion.Oracle7
						comboSqlDialect.SelectedItem = "Oracle 7"
						Exit Select
					Case OracleServerVersion.Oracle8
						comboSqlDialect.SelectedItem = "Oracle 8"
						Exit Select
					Case OracleServerVersion.Oracle9
						comboSqlDialect.SelectedItem = "Oracle 9"
						Exit Select
					Case OracleServerVersion.Oracle10
						comboSqlDialect.SelectedItem = "Oracle 10"
						Exit Select
					Case OracleServerVersion.Oracle11
						comboSqlDialect.SelectedItem = "Oracle 11g"
						Exit Select
					Case OracleServerVersion.Oracle12
						comboSqlDialect.SelectedItem = "Oracle 12c"
						Exit Select
					Case OracleServerVersion.Oracle18
						comboSqlDialect.SelectedItem = "Oracle 18c"
						Exit Select
					Case OracleServerVersion.Oracle19
						comboSqlDialect.SelectedItem = "Oracle 19c"
						Exit Select
					Case Else
						comboSqlDialect.SelectedItem = "Oracle 18c"
						Exit Select
				End Select
			ElseIf TypeOf _sqlContext.SyntaxProvider Is PostgreSQLSyntaxProvider Then
				comboSqlDialect.SelectedItem = "PostgreSQL"
			ElseIf TypeOf _sqlContext.SyntaxProvider Is SQLiteSyntaxProvider Then
				comboSqlDialect.SelectedItem = "SQLite"
			ElseIf TypeOf _sqlContext.SyntaxProvider Is SybaseSyntaxProvider Then
				Select Case TryCast(_sqlContext.SyntaxProvider, SybaseSyntaxProvider).ServerVersion
					Case SybaseServerVersion.SybaseASE
						comboSqlDialect.SelectedItem = "Sybase ASE"
						Exit Select
					Case SybaseServerVersion.SybaseASA
						comboSqlDialect.SelectedItem = "Sybase SQL Anywhere"
						Exit Select
					Case Else
						comboSqlDialect.SelectedItem = "Sybase SQL Anywhere"
						Exit Select
				End Select
			ElseIf TypeOf _sqlContext.SyntaxProvider Is TeradataSyntaxProvider Then
				comboSqlDialect.SelectedItem = "Teradata"
			ElseIf TypeOf _sqlContext.SyntaxProvider Is VistaDBSyntaxProvider Then
				comboSqlDialect.SelectedItem = "VistaDB"
			End If

			If TypeOf _sqlContext.SyntaxProvider Is GenericSyntaxProvider Then
				comboSqlDialect.SelectedItem = "Generic"
			End If

			If _sqlContext.SyntaxProvider IsNot Nothing Then
				comboIdentCaseSens.SelectedIndex = CInt(_sqlContext.SyntaxProvider.IdentCaseSens)
				textBeginQuotationSymbol.Text = _sqlContext.SyntaxProvider.QuoteBegin
				textEndQuotationSymbol.Text = _sqlContext.SyntaxProvider.QuoteEnd
			End If

			cbQuoteAllIdentifiers.IsChecked = _sqlContext.SQLGenerationOptionsForServer.QuoteIdentifiers = IdentQuotation.All

			AddHandler comboSqlDialect.SelectionChanged, AddressOf comboSqlDialect_SelectionChanged
			AddHandler comboIdentCaseSens.SelectionChanged, AddressOf comboIdentCaseSens_SelectionChanged
			AddHandler cbQuoteAllIdentifiers.Checked, AddressOf cbQuoteAllIdentifiers_Checked
			AddHandler cbQuoteAllIdentifiers.Unchecked, AddressOf cbQuoteAllIdentifiers_Checked
		End Sub

		Private Sub cbQuoteAllIdentifiers_Checked(sender As Object, e As System.Windows.RoutedEventArgs)
			Modified = True
		End Sub

		Private Sub comboIdentCaseSens_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
			_syntaxProvider.IdentCaseSens = CType(comboIdentCaseSens.SelectedIndex, IdentCaseSensitivity)
			comboIdentCaseSens.SelectedIndex = CInt(_syntaxProvider.IdentCaseSens)

			Modified = True
		End Sub

		Private Sub comboSqlDialect_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
			Select Case comboSqlDialect.SelectedItem.ToString()
				Case "ANSI SQL-92"
					_syntaxProvider = New SQL92SyntaxProvider()
					Exit Select
				Case "Auto"
					_syntaxProvider = New AutoSyntaxProvider()
					Exit Select
				Case "ANSI SQL-89"
					_syntaxProvider = New SQL89SyntaxProvider()
					Exit Select
				Case "ANSI SQL-2003"
					_syntaxProvider = New SQL2003SyntaxProvider()
					Exit Select
				Case "Firebird 1.0"
					_syntaxProvider = New FirebirdSyntaxProvider()
					TryCast(_syntaxProvider, FirebirdSyntaxProvider).ServerVersion = FirebirdVersion.Firebird10
					Exit Select
				Case "Firebird 1.5"
					_syntaxProvider = New FirebirdSyntaxProvider()
					TryCast(_syntaxProvider, FirebirdSyntaxProvider).ServerVersion = FirebirdVersion.Firebird15
					Exit Select
				Case "Firebird 2.0"
					_syntaxProvider = New FirebirdSyntaxProvider()
					TryCast(_syntaxProvider, FirebirdSyntaxProvider).ServerVersion = FirebirdVersion.Firebird20
					Exit Select
				Case "Firebird 2.5"
					_syntaxProvider = New FirebirdSyntaxProvider()
					TryCast(_syntaxProvider, FirebirdSyntaxProvider).ServerVersion = FirebirdVersion.Firebird25
					Exit Select
				Case "IBM DB2"
					_syntaxProvider = New DB2SyntaxProvider()
					Exit Select
				Case "IBM Informix 8"
					_syntaxProvider = New InformixSyntaxProvider()
					TryCast(_syntaxProvider, InformixSyntaxProvider).ServerVersion = InformixVersion.DS8
					Exit Select
				Case "IBM Informix 9"
					_syntaxProvider = New InformixSyntaxProvider()
					TryCast(_syntaxProvider, InformixSyntaxProvider).ServerVersion = InformixVersion.DS9
					Exit Select
				Case "IBM Informix 10"
					_syntaxProvider = New InformixSyntaxProvider()
					TryCast(_syntaxProvider, InformixSyntaxProvider).ServerVersion = InformixVersion.DS10
					Exit Select
				Case "MS Access 97 (MS Jet 3.0)"
					_syntaxProvider = New MSAccessSyntaxProvider()
					TryCast(_syntaxProvider, MSAccessSyntaxProvider).ServerVersion = MSAccessServerVersion.MSJET3
					Exit Select
				Case "MS Access 2000 (MS Jet 4.0)", "MS Access XP (MS Jet 4.0)", "MS Access 2003 (MS Jet 4.0)"
					_syntaxProvider = New MSAccessSyntaxProvider()
					TryCast(_syntaxProvider, MSAccessSyntaxProvider).ServerVersion = MSAccessServerVersion.MSJET4
					Exit Select
				Case "MS SQL Server 7"
					_syntaxProvider = New MSSQLSyntaxProvider()
					TryCast(_syntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL7
					Exit Select
				Case "MS SQL Server 2000"
					_syntaxProvider = New MSSQLSyntaxProvider()
					TryCast(_syntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2000
					Exit Select
				Case "MS SQL Server 2005"
					_syntaxProvider = New MSSQLSyntaxProvider()
					TryCast(_syntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2005
					Exit Select
				Case "MS SQL Server 2008"
					_syntaxProvider = New MSSQLSyntaxProvider()
					TryCast(_syntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2008
					Exit Select
				Case "MS SQL Server 2012"
					_syntaxProvider = New MSSQLSyntaxProvider()
					TryCast(_syntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2012
					Exit Select
				Case "MS SQL Server 2014"
					_syntaxProvider = New MSSQLSyntaxProvider()
					TryCast(_syntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2014
					Exit Select
				Case "MS SQL Server 2016"
					_syntaxProvider = New MSSQLSyntaxProvider()
					TryCast(_syntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2016
					Exit Select
				Case "MS SQL Server 2017"
					_syntaxProvider = New MSSQLSyntaxProvider()
					TryCast(_syntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2017
					Exit Select
				Case "MS SQL Server 2019"
					_syntaxProvider = New MSSQLSyntaxProvider()
					TryCast(_syntaxProvider, MSSQLSyntaxProvider).ServerVersion = MSSQLServerVersion.MSSQL2019
					Exit Select
				Case "MS SQL Server Compact Edition"
					_syntaxProvider = New MSSQLCESyntaxProvider()
					Exit Select
				Case "MySQL 3.xx"
					_syntaxProvider = New MySQLSyntaxProvider()
					TryCast(_syntaxProvider, MySQLSyntaxProvider).ServerVersionInt = 39999
					Exit Select
				Case "MySQL 4.0"
					_syntaxProvider = New MySQLSyntaxProvider()
					TryCast(_syntaxProvider, MySQLSyntaxProvider).ServerVersionInt = 40099
					Exit Select
				Case "MySQL 4.1"
					_syntaxProvider = New MySQLSyntaxProvider()
					TryCast(_syntaxProvider, MySQLSyntaxProvider).ServerVersionInt = 49999
					Exit Select
				Case "MySQL 5.0"
					_syntaxProvider = New MySQLSyntaxProvider()
					TryCast(_syntaxProvider, MySQLSyntaxProvider).ServerVersionInt = 50000
					Exit Select
				Case "MySQL 8.0"
					_syntaxProvider = New MySQLSyntaxProvider()
					TryCast(_syntaxProvider, MySQLSyntaxProvider).ServerVersionInt = 80012
					Exit Select
				Case "Oracle 7"
					_syntaxProvider = New OracleSyntaxProvider()
					TryCast(_syntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle7
					Exit Select
				Case "Oracle 8"
					_syntaxProvider = New OracleSyntaxProvider()
					TryCast(_syntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle8
					Exit Select
				Case "Oracle 9"
					_syntaxProvider = New OracleSyntaxProvider()
					TryCast(_syntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle9
					Exit Select
				Case "Oracle 10"
					_syntaxProvider = New OracleSyntaxProvider()
					TryCast(_syntaxProvider, OracleSyntaxProvider).ServerVersion = OracleServerVersion.Oracle10
					Exit Select
				Case "Oracle 11g"
					_syntaxProvider = New OracleSyntaxProvider() With { _
						.ServerVersion = OracleServerVersion.Oracle11 _
					}
					Exit Select
				Case "Oracle 12c"
					_syntaxProvider = New OracleSyntaxProvider() With { _
						.ServerVersion = OracleServerVersion.Oracle12 _
					}
					Exit Select
				Case "Oracle 18c"
					_syntaxProvider = New OracleSyntaxProvider() With { _
						.ServerVersion = OracleServerVersion.Oracle18 _
					}
					Exit Select
				Case "Oracle 19c"
					_syntaxProvider = New OracleSyntaxProvider() With { _
						.ServerVersion = OracleServerVersion.Oracle19 _
					}
					Exit Select
				Case "PostgreSQL"
					_syntaxProvider = New PostgreSQLSyntaxProvider()
					Exit Select
				Case "SQLite"
					_syntaxProvider = New SQLiteSyntaxProvider()
					Exit Select
				Case "Sybase ASE"
					_syntaxProvider = New SybaseSyntaxProvider()
					TryCast(_syntaxProvider, SybaseSyntaxProvider).ServerVersion = SybaseServerVersion.SybaseASE
					Exit Select
				Case "Sybase SQL Anywhere"
					_syntaxProvider = New SybaseSyntaxProvider()
					TryCast(_syntaxProvider, SybaseSyntaxProvider).ServerVersion = SybaseServerVersion.SybaseASA
					Exit Select
				Case "Teradata"
					_syntaxProvider = New TeradataSyntaxProvider()
					Exit Select
				Case "VistaDB"
					_syntaxProvider = New VistaDBSyntaxProvider()
					Exit Select
				Case "Generic"
					_syntaxProvider = New GenericSyntaxProvider()
					DirectCast(_syntaxProvider, GenericSyntaxProvider).RedetectServer(_sqlContext)
					Exit Select
				Case Else
					_syntaxProvider = New GenericSyntaxProvider()
					Exit Select
			End Select

			comboIdentCaseSens.SelectedIndex = CInt(_syntaxProvider.IdentCaseSens)
			textBeginQuotationSymbol.Text = _syntaxProvider.QuoteBegin
			textEndQuotationSymbol.Text = _syntaxProvider.QuoteEnd

			Modified = True
		End Sub

		Public Sub ApplyChanges()
			If Not Modified Then
				Return
			End If

            Dim oldSyntaxProvider As BaseSyntaxProvider = _sqlContext.SyntaxProvider

            _sqlContext.SyntaxProvider = _syntaxProvider

			_sqlContext.SQLGenerationOptionsForServer.QuoteIdentifiers = If(cbQuoteAllIdentifiers.IsChecked.HasValue AndAlso cbQuoteAllIdentifiers.IsChecked.Value, IdentQuotation.All, IdentQuotation.IfNeed)

			If oldSyntaxProvider IsNot Nothing Then
				oldSyntaxProvider.Dispose()
			End If
		End Sub
	End Class
End Namespace
