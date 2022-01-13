''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2022 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System
Imports System.ComponentModel
Imports ActiveQueryBuilder.View.WPF

Namespace QueryBuilderProperties
    ''' <summary>
    ''' Interaction logic for PanesVisibilityPage.xaml
    ''' </summary>
    <ToolboxItem(False)>
    Partial Public Class PanesVisibilityPage
        Private ReadOnly _queryBuilder As QueryBuilder
        Public Property Modified() As Boolean

        Public Sub New(qb As QueryBuilder)
            Modified = False
            _queryBuilder = qb

            InitializeComponent()

            cbShowDesignPane.IsChecked = _queryBuilder.PanesConfigurationOptions.DesignPaneVisible
            cbShowQueryColumnsPane.IsChecked = _queryBuilder.PanesConfigurationOptions.QueryColumnsPaneVisible
            cbShowDatabaseSchemaView.IsChecked = _queryBuilder.PanesConfigurationOptions.DatabaseSchemaViewVisible
            cbShowQueryNavigationBar.IsChecked = _queryBuilder.PanesConfigurationOptions.QueryNavigationBarVisible

            AddHandler cbShowDesignPane.Checked, AddressOf Changed
            AddHandler cbShowDesignPane.Unchecked, AddressOf Changed
            AddHandler cbShowQueryColumnsPane.Checked, AddressOf Changed
            AddHandler cbShowQueryColumnsPane.Unchecked, AddressOf Changed
            AddHandler cbShowDatabaseSchemaView.Checked, AddressOf Changed
            AddHandler cbShowDatabaseSchemaView.Unchecked, AddressOf Changed
            AddHandler cbShowQueryNavigationBar.Checked, AddressOf Changed
            AddHandler cbShowQueryNavigationBar.Unchecked, AddressOf Changed
        End Sub

        Public Sub New()
            Modified = False
            InitializeComponent()
        End Sub

        Private Sub Changed(sender As Object, e As EventArgs)
            If Equals(sender, cbShowDesignPane) Then
                If Not (cbShowDesignPane.IsChecked.HasValue AndAlso cbShowDesignPane.IsChecked.Value) AndAlso Not (cbShowQueryColumnsPane.IsChecked.HasValue AndAlso cbShowQueryColumnsPane.IsChecked.Value) Then
                    cbShowQueryColumnsPane.IsChecked = True
                End If
            ElseIf Equals(sender, cbShowQueryColumnsPane) Then
                If Not (cbShowDesignPane.IsChecked.HasValue AndAlso cbShowDesignPane.IsChecked.Value) AndAlso Not (cbShowQueryColumnsPane.IsChecked.HasValue AndAlso cbShowQueryColumnsPane.IsChecked.Value) Then
                    cbShowDesignPane.IsChecked = True
                End If
            End If

            Modified = True
        End Sub

        Public Sub ApplyChanges()
            If Modified Then
                _queryBuilder.PanesConfigurationOptions.BeginUpdate()

                Try
                    _queryBuilder.PanesConfigurationOptions.DesignPaneVisible = cbShowDesignPane.IsChecked.HasValue AndAlso cbShowDesignPane.IsChecked.Value
                    _queryBuilder.PanesConfigurationOptions.QueryColumnsPaneVisible = cbShowQueryColumnsPane.IsChecked.HasValue AndAlso cbShowQueryColumnsPane.IsChecked.Value
                    _queryBuilder.PanesConfigurationOptions.DatabaseSchemaViewVisible = cbShowDatabaseSchemaView.IsChecked.HasValue AndAlso cbShowDatabaseSchemaView.IsChecked.Value
                    _queryBuilder.PanesConfigurationOptions.QueryNavigationBarVisible = cbShowQueryNavigationBar.IsChecked.HasValue AndAlso cbShowQueryNavigationBar.IsChecked.Value
                Finally
                    _queryBuilder.PanesConfigurationOptions.EndUpdate()
                End Try
            End If
        End Sub
    End Class
End Namespace
