''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2025 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input

Namespace ConnectionWindows
    ''' <summary>
    ''' Interaction logic for MSSQLConnectionWindow.xaml
    ''' </summary>
    Public Partial Class MSSQLConnectionWindow
        Public ConnectionString As String = ""

        Public Sub New()
            InitializeComponent()

            AddHandler Loaded, AddressOf MSSQLConnectionWindow_Loaded
        End Sub

        Private Sub MSSQLConnectionWindow_Loaded(sender As Object, e As RoutedEventArgs)
            comboBoxAuthentication.SelectedIndex = 0
        End Sub

        Private Sub ComboBoxAuthentication_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            If comboBoxAuthentication.SelectedIndex = 0 Then
                'Windows Authentication
                textBoxLogin.IsEnabled = False
                textBoxPassword.IsEnabled = False
            Else
                'SQL Server Authentication
                textBoxLogin.IsEnabled = True
                textBoxPassword.IsEnabled = True
            End If
        End Sub

        Private Sub ComboBoxDatabase_OnDropDownOpened(sender As Object, e As EventArgs)
            ' Fill the drop down list with available database names

            Dim builder As New SqlConnectionStringBuilder()

            builder.DataSource = textBoxServerName.Text

            If comboBoxAuthentication.SelectedIndex = 0 Then
                builder.IntegratedSecurity = True
            Else
                builder.IntegratedSecurity = False
                builder.UserID = textBoxLogin.Text
                builder.Password = textBoxPassword.Text
            End If

            ' try to connect
            Using connection As New SqlConnection(builder.ConnectionString)
                Mouse.OverrideCursor = Cursors.Wait

                Dim currentDatabase As String = comboBoxDatabase.SelectedItem.ToString()

                comboBoxDatabase.Items.Clear()
                comboBoxDatabase.Items.Add("<default>")
                comboBoxDatabase.SelectedIndex = 0

                Try
                    connection.Open()

                    ' connected successfully
                    ' retrieve available databases

                    Dim schemaTable As DataTable = connection.GetSchema("Databases")

                    For Each r As DataRow In schemaTable.Rows
                        comboBoxDatabase.Items.Add(r(0))
                    Next

                    comboBoxDatabase.SelectedItem = currentDatabase
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Connection Failure.")
                Finally
                    Mouse.OverrideCursor = Nothing
                End Try
            End Using
        End Sub

        Private Sub buttonConnect_Click(sender As Object, e As RoutedEventArgs)
            Dim builder As New SqlConnectionStringBuilder()

            builder.DataSource = textBoxServerName.Text

            If comboBoxAuthentication.SelectedIndex = 0 Then
                builder.IntegratedSecurity = True
            Else
                builder.IntegratedSecurity = False
                builder.UserID = textBoxLogin.Text
                builder.Password = textBoxPassword.Text
            End If

            If comboBoxDatabase.SelectedIndex > 0 Then
                builder.InitialCatalog = comboBoxDatabase.SelectedItem.ToString()
            End If

            ' check the connection

            Using connection As New SqlConnection(builder.ConnectionString)
                Mouse.OverrideCursor = Cursors.Wait

                Try
                    connection.Open()
                    ConnectionString = builder.ConnectionString
                    DialogResult = True
                Catch ex As Exception

                    MessageBox.Show(ex.Message, "Connection Failure.")
                Finally
                    Mouse.OverrideCursor = Nothing
                End Try
            End Using
        End Sub

        Private Sub ButtonCancel_OnClick(sender As Object, e As RoutedEventArgs)
            DialogResult = False
        End Sub
    End Class
End Namespace
