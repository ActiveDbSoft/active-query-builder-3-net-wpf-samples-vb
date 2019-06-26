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
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.Windows
Imports System.Windows.Controls
Imports ActiveQueryBuilder.Core

Partial Friend Class ErrorBox
    Public ReadOnly SyntaxProviders As ObservableCollection(Of BaseSyntaxProvider) = New ObservableCollection(Of BaseSyntaxProvider)()

    Public Event SyntaxProviderChanged As SelectionChangedEventHandler
    public event  GoToErrorPositionEvent as EventHandler
    public event  RevertValidTextEvent as EventHandler

    Public Shared ReadOnly MessageProperty As DependencyProperty = DependencyProperty.Register("Message", GetType(String), GetType(ErrorBox),  New FrameworkPropertyMetadata(Nothing, AddressOf MessageChanged))
    Private _allowChangedSyntax As Boolean = True

    Public Property Message As String
        Get
            Return DirectCast(GetValue(MessageProperty), String)
        End Get
        Set
            SetValue(MessageProperty, Value)
        End Set
    End Property

    Private Shared Sub MessageChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
        Dim self As ErrorBox = CType(d, ErrorBox)
        Dim value As String = CType(e.NewValue, String)

        self.TextBlockErrorPrompt.Text = value
        self.Visibility = If(String.IsNullOrEmpty(value), Visibility.Collapsed, Visibility.Visible)
    End Sub
    Public Sub New()
        InitializeComponent()
        Visibility = Visibility.Collapsed
       
        AddHandler SyntaxProviders.CollectionChanged, AddressOf CollectionChanged
    End Sub

    Private Sub CollectionChanged(sender As Object, e As NotifyCollectionChangedEventArgs)
        _allowChangedSyntax = false
        ComboBoxSyntaxProvider.Items.Clear()
        For Each baseSyntaxProvider As BaseSyntaxProvider In SyntaxProviders
            ComboBoxSyntaxProvider.Items.Add(new ComboBoxItem(baseSyntaxProvider))    
        Next
        
        _allowChangedSyntax = true
    End Sub

    Private Sub ComboBoxSyntaxProvider_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If Not _allowChangedSyntax Then Return
        Dim syntaxProvider = CType(ComboBoxSyntaxProvider.SelectedItem, ComboBoxItem).SyntaxProvider
        OnSyntaxProviderChanged(New SelectionChangedEventArgs(e.RoutedEvent, New List(Of BaseSyntaxProvider)(), New List(Of BaseSyntaxProvider) From {
                                                                 syntaxProvider
                                                                 }))
    End Sub

    Public Sub SetCurrentSyntaxProvider(syntaxProvider As BaseSyntaxProvider)
        _allowChangedSyntax = False
        ComboBoxSyntaxProvider.Text = syntaxProvider.ToString()
        _allowChangedSyntax = True
    End Sub

    Protected Overridable Sub OnSyntaxProviderChanged(e As SelectionChangedEventArgs)
        RaiseEvent SyntaxProviderChanged(Me, e)
    End Sub

    Private Sub HyperlinkGoToPosition_OnClick(sender As Object, e As RoutedEventArgs)
        RaiseEvent GoToErrorPositionEvent(Me, e)
    End Sub

    Private Sub HyperlinkPreviousValidText_OnClick(sender As Object, e As RoutedEventArgs)
        RaiseEvent RevertValidTextEvent(Me, e)
    End Sub
End Class

Public Class ComboBoxItem
    Public ReadOnly Property SyntaxProvider As BaseSyntaxProvider

    Public ReadOnly Property DisplayString As String
        Get
            Return SyntaxProvider.ToString()
        End Get
    End Property

    Public Sub New()
    End Sub

    Public Sub New(ByVal provider As BaseSyntaxProvider)
        SyntaxProvider = provider
    End Sub
End Class
