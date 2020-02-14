'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System
Imports System.Windows
Imports System.Windows.Controls
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.Core.PropertiesEditors
Imports ActiveQueryBuilder.View
Imports ActiveQueryBuilder.View.PropertiesEditors
Imports Microsoft.Win32
Imports Helpers = ActiveQueryBuilder.Core.Helpers

Partial Public Class XmlConnectionEditWindow
    Inherits Window

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
        Dim syntxProps As ObjectProperties = _connection.ConnectionDescriptor.SyntaxProperties

        If syntxProps Is Nothing Then
            pbSyntax.Visibility = Visibility.Collapsed
            Return
        End If

        pbSyntax.Visibility = Visibility.Visible
        ClearProperties(syntxProps)
        Dim container As IPropertiesContainer = PropertiesFactory.GetPropertiesContainer(syntxProps)
        TryCast(pbSyntax, IPropertiesControl).SetProperties(container)
        End Sub

    Private Sub ClearProperties(properties As ObjectProperties)
        properties.GroupProperties.Clear()
        properties.PropertiesEditors.Clear()
    End Sub

    Private Sub FillSyntaxTypes()
        For Each syntax As Type In Helpers.SyntaxProviderList
            Dim instance As BaseSyntaxProvider = TryCast(Activator.CreateInstance(syntax), BaseSyntaxProvider)
            cbSyntax.Items.Add(instance.Description)
        Next
    End Sub

    Private Sub ButtonOk_OnClick(sender As Object, e As RoutedEventArgs)
        DialogResult = True
    End Sub

    Private Sub ButtonCancel_OnClick(sender As Object, e As RoutedEventArgs)
        DialogResult = False
    End Sub

    Private Sub CbSyntax_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Dim syntaxType As Type = GetSelectedSyntaxType()

        If _connection.ConnectionDescriptor.SyntaxProvider.[GetType]() = syntaxType Then
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
        Dim openFileDialog As OpenFileDialog = New OpenFileDialog With {
            .Filter = "XML files|*xml|All files|*.*"
        }

        If openFileDialog.ShowDialog() = True Then
            tbXmlPath.Text = openFileDialog.FileName
        End If
    End Sub
End Class
