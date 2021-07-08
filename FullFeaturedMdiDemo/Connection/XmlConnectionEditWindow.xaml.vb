''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2021 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports ActiveQueryBuilder.Core.PropertiesEditors
Imports ActiveQueryBuilder.View.PropertiesEditors
Imports GeneralAssembly
Imports Microsoft.Win32
Imports Helpers = ActiveQueryBuilder.Core.Helpers

Namespace Connection
    ''' <summary>
    ''' Interaction logic for XmlConnectionEditWindow.xaml
    ''' </summary>
    Partial Public Class XmlConnectionEditWindow
        Private ReadOnly _connection As ConnectionInfo
        Public Sub New()
            InitializeComponent()
        End Sub

        Public Sub New(connection As ConnectionInfo)
            Me.New()
            _connection = connection
            FillSyntaxTypes()

            tbConnectionName.Text = _connection.Name
            tbXmlPath.Text = _connection.XMLPath
            cbSyntax.SelectedItem = _connection.ConnectionDescriptor.SyntaxProvider.Description

            RecreateSyntaxFrame()
        End Sub

        Private Function GetSelectedSyntaxType() As Type
            Return Helpers.SyntaxProviderList(cbSyntax.SelectedIndex)
        End Function

        Private Function CreateSyntaxProvider(type As Type) As BaseSyntaxProvider
            Return TryCast(Activator.CreateInstance(type), BaseSyntaxProvider)
        End Function

        Private Sub RecreateSyntaxFrame()
            pbSyntax.ClearProperties()
            Dim syntxProps = _connection.ConnectionDescriptor.SyntaxProperties
            If syntxProps Is Nothing Then
                pbSyntax.Visibility = Visibility.Collapsed
                Return
            End If

            pbSyntax.Visibility = Visibility.Visible
            ClearProperties(syntxProps)
            Dim container = PropertiesFactory.GetPropertiesContainer(syntxProps)
            TryCast(pbSyntax, IPropertiesControl).SetProperties(container)
        End Sub

        Private Sub ClearProperties(properties As ObjectProperties)
            properties.GroupProperties.Clear()
            properties.PropertiesEditors.Clear()
        End Sub

        Private Sub FillSyntaxTypes()
            For Each syntax As Type In Helpers.SyntaxProviderList
                Dim instance = TryCast(Activator.CreateInstance(syntax), BaseSyntaxProvider)
                cbSyntax.Items.Add(instance.Description)
            Next syntax
        End Sub

        Private Sub ButtonOk_OnClick(sender As Object, e As RoutedEventArgs)
            DialogResult = True
        End Sub

        Private Sub ButtonCancel_OnClick(sender As Object, e As RoutedEventArgs)
            DialogResult = False
        End Sub

        Private Sub CbSyntax_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            Dim syntaxType = GetSelectedSyntaxType()
            If _connection.ConnectionDescriptor.SyntaxProvider.GetType() Is syntaxType Then
                Return
            End If

            _connection.ConnectionDescriptor.SyntaxProvider = CreateSyntaxProvider(syntaxType)
            _connection.SyntaxProviderName = syntaxType.ToString()
            RecreateSyntaxFrame()
        End Sub

        Private Sub TbConnectionName_OnTextChanged(sender As Object, e As TextChangedEventArgs)
            _connection.Name = tbConnectionName.Text
        End Sub

        Private Sub TbXmlPath_OnTextChanged(sender As Object, e As TextChangedEventArgs)
            _connection.XMLPath = tbXmlPath.Text
        End Sub

        Private Sub ButtonBase_OnClick(sender As Object, e As RoutedEventArgs)
            Dim openFileDialog = New OpenFileDialog With {.Filter = "XML files|*xml|All files|*.*"}

            If openFileDialog.ShowDialog().Equals(True) Then
                tbXmlPath.Text = openFileDialog.FileName
            End If
        End Sub
    End Class
End Namespace
