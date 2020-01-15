'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports ActiveQueryBuilder.Core.PropertiesEditors
Imports ActiveQueryBuilder.View
Imports ActiveQueryBuilder.View.PropertiesEditors
Imports ActiveQueryBuilder.View.WPF.DatabaseSchemaView
Imports ActiveQueryBuilder.View.WPF.ExpressionEditor
Imports ActiveQueryBuilder.View.WPF.QueryView
Imports FullFeaturedMdiDemo.PropertiesForm.Tabs

Namespace PropertiesForm
    Public Partial Class QueryPropertiesWindow
        Private ReadOnly _linkToPageGeneral As Dictionary(Of TextBlock, Grid) = New Dictionary(Of TextBlock, Grid)()
        Private ReadOnly _linkToPageFormatting As Dictionary(Of TextBlock, UserControl) = New Dictionary(Of TextBlock, UserControl)()
        Private ReadOnly _sqlGenerationControl As UserControl
        Private ReadOnly _textEditorOptions As TextEditorOptions = New TextEditorOptions()
        Private ReadOnly _textEditorSqlOptions As SqlTextEditorOptions = New SqlTextEditorOptions()
        Private ReadOnly _childWindow As ChildWindow
        Private ReadOnly _dbSchemaView As DatabaseSchemaView
        Private _currentGeneralSelectedLink As TextBlock
        Private _currentFormattingSelectedLink As TextBlock
        Private _structureOptionsChanged As Boolean

        Public Sub New()
            InitializeComponent()
        End Sub

        Public Sub New(childWindow As ChildWindow, schemaView As DatabaseSchemaView)
            InitializeComponent()
            _childWindow = childWindow
            _dbSchemaView = schemaView
            linkAddObject.Visibility = Visibility.Collapsed
            _linkToPageFormatting.Add(LinkMain, New MainQueryTab(childWindow.SqlFormattingOptions))
            _linkToPageFormatting.Add(LinkMainCommon, New CommonTab(childWindow.SqlFormattingOptions, childWindow.SqlFormattingOptions.MainQueryFormat))
            _linkToPageFormatting.Add(LinkMainExpressions, New ExpressionsTab(childWindow.SqlFormattingOptions, childWindow.SqlFormattingOptions.MainQueryFormat))
            _linkToPageFormatting.Add(LinkCte, New SubQueryTab(childWindow.SqlFormattingOptions, SubQueryType.Cte))
            _linkToPageFormatting.Add(LinkCteCommon, New CommonTab(childWindow.SqlFormattingOptions, childWindow.SqlFormattingOptions.CTESubQueryFormat))
            _linkToPageFormatting.Add(LinkCteExpressions, New ExpressionsTab(childWindow.SqlFormattingOptions, childWindow.SqlFormattingOptions.CTESubQueryFormat))
            _linkToPageFormatting.Add(LinkDerived, New SubQueryTab(childWindow.SqlFormattingOptions, SubQueryType.Derived))
            _linkToPageFormatting.Add(LinkDerivedCommon, New CommonTab(childWindow.SqlFormattingOptions, childWindow.SqlFormattingOptions.DerivedQueryFormat))
            _linkToPageFormatting.Add(LinkDerivedExpressions, New ExpressionsTab(childWindow.SqlFormattingOptions, childWindow.SqlFormattingOptions.DerivedQueryFormat))
            _linkToPageFormatting.Add(LinkExpression, New SubQueryTab(childWindow.SqlFormattingOptions, SubQueryType.Expression))
            _linkToPageFormatting.Add(LinkExpressionCommon, New CommonTab(childWindow.SqlFormattingOptions, childWindow.SqlFormattingOptions.ExpressionSubQueryFormat))
            _linkToPageFormatting.Add(LinkExpressionExpressions, New ExpressionsTab(childWindow.SqlFormattingOptions, childWindow.SqlFormattingOptions.ExpressionSubQueryFormat))
            _sqlGenerationControl = New SqlGenerationPage(childWindow.SqlGenerationOptions, childWindow.SqlFormattingOptions)
            _linkToPageGeneral.Add(linkBehavior, GetPropertyPage(New ObjectProperties(childWindow.BehaviorOptions)))
            _linkToPageGeneral.Add(linkSchemaView, GetPropertyPage(New ObjectProperties(schemaView.Options)))
            _linkToPageGeneral.Add(linkDesignPane, GetPropertyPage(New ObjectProperties(childWindow.DesignPaneOptions)))
            _linkToPageGeneral.Add(linkVisual, GetPropertyPage(New ObjectProperties(childWindow.VisualOptions)))
            _linkToPageGeneral.Add(linkDatasource, GetPropertyPage(New ObjectProperties(childWindow.DataSourceOptions)))
            _linkToPageGeneral.Add(linkMetadataLoading, GetPropertyPage(New ObjectProperties(childWindow.MetadataLoadingOptions)))
            _linkToPageGeneral.Add(linkMetadataStructure, GetPropertyPage(New ObjectProperties(childWindow.MetadataStructureOptions)))
            _linkToPageGeneral.Add(linkQueryColumnList, GetPropertyPage(New ObjectProperties(childWindow.QueryColumnListOptions)))
            _linkToPageGeneral.Add(linkQueryNavBar, GetPropertyPage(New ObjectProperties(childWindow.QueryNavBarOptions)))
            _linkToPageGeneral.Add(linkUserInterface, GetPropertyPage(New ObjectProperties(childWindow.UserInterfaceOptions)))
            _linkToPageGeneral.Add(linkExpressionEditor, GetPropertyPage(New ObjectProperties(childWindow.ExpressionEditorOptions)))
            _textEditorOptions.Assign(childWindow.TextEditorOptions)
            AddHandler _textEditorOptions.Updated, AddressOf TextEditorOptionsOnUpdated
            _linkToPageGeneral.Add(linkTextEditor, GetPropertyPage(New ObjectProperties(_textEditorOptions)))
            _textEditorSqlOptions.Assign(childWindow.TextEditorSqlOptions)
            AddHandler _textEditorSqlOptions.Updated, AddressOf TextEditorOptionsOnUpdated
            _linkToPageGeneral.Add(linkTextEditorSql, GetPropertyPage(New ObjectProperties(_textEditorSqlOptions)))
            GeneralLinkClick(linkGeneration, Nothing)
            FormattingLinkClick(LinkMain, Nothing)
            AddHandler childWindow.MetadataStructureOptions.Updated, AddressOf MetadataStructureOptionsOnUpdated
        End Sub

        Protected Overrides Sub OnClosing(e As CancelEventArgs)
            Properties.Settings.Default.Options = _childWindow.GetOptions().SerializeToString()
            Properties.Settings.Default.Save()
            If _structureOptionsChanged Then _dbSchemaView.InitializeDatabaseSchemaTree()
            MyBase.OnClosing(e)
        End Sub

        Private Sub MetadataStructureOptionsOnUpdated(sender As Object, e As EventArgs)
            _structureOptionsChanged = True
        End Sub

        Private Sub TextEditorOptionsOnUpdated(sender As Object, eventArgs As EventArgs)
            _childWindow.TextEditorOptions = _textEditorOptions
            _childWindow.TextEditorSqlOptions = _textEditorSqlOptions
        End Sub

        Private Function GetPropertyPage(propertiesObject As ObjectProperties) As Grid
            Dim propertiesContainer As IPropertiesContainer = PropertiesFactory.GetPropertiesContainer(propertiesObject)
            Dim propertyPage As PropertiesBar = New PropertiesBar()
            Dim propertiesControl As IPropertiesControl = CType(propertyPage, IPropertiesControl)
            propertiesControl.SetProperties(propertiesContainer)
            Return propertyPage
        End Function

        Private Sub GeneralLinkClick(sender As Object, e As MouseButtonEventArgs)
            If _currentGeneralSelectedLink IsNot Nothing Then _currentGeneralSelectedLink.Foreground = Brushes.Black
            _currentGeneralSelectedLink = CType(sender, TextBlock)
            _currentGeneralSelectedLink.Foreground = Brushes.Blue

            If Equals(_currentGeneralSelectedLink, linkGeneration) Then
                gridGeneral.Children.Clear()
                _sqlGenerationControl.Margin = New Thickness(10, 10, 0, 0)
                gridGeneral.Children.Add(_sqlGenerationControl)
                Return
            End If

            SwitchGeneralPage(_linkToPageGeneral(_currentGeneralSelectedLink))
        End Sub

        Private Sub SwitchGeneralPage(page As Grid)
            gridGeneral.Children.Clear()
            page.Margin = New Thickness(10, 10, 0, 0)
            gridGeneral.Children.Add(page)
        End Sub

        Private Sub SwitchFormattingPage(page As UserControl)
            gridFormatting.Children.Clear()
            page.Margin = New Thickness(10, 10, 0, 0)
            gridFormatting.Children.Add(page)
        End Sub

        Private Sub ButtonCancel_OnClick(sender As Object, e As RoutedEventArgs)
            DialogResult = True
        End Sub

        Private Sub FormattingLinkClick(sender As Object, e As MouseButtonEventArgs)
            If _currentFormattingSelectedLink IsNot Nothing Then _currentFormattingSelectedLink.Foreground = Brushes.Black
            _currentFormattingSelectedLink = CType(sender, TextBlock)
            _currentFormattingSelectedLink.Foreground = Brushes.Blue
            SwitchFormattingPage(_linkToPageFormatting(_currentFormattingSelectedLink))
        End Sub
    End Class
End Namespace
