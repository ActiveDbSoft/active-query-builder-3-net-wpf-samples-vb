''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2023 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Windows
Imports System.Windows.Controls
Imports ActiveQueryBuilder.Core

Namespace Common
    Partial Public Class SqlErrorBox
        Private _allowChangedSyntax As Boolean = True

        Public Event SyntaxProviderChanged As SelectionChangedEventHandler
        Public Event GoToErrorPosition As EventHandler
        Public Event RevertValidText As EventHandler

        Public Shared ReadOnly VisibilityCheckSyntaxBlockProperty As DependencyProperty = DependencyProperty.Register("VisibilityCheckSyntaxBlock", GetType(Visibility), GetType(SqlErrorBox), New PropertyMetadata(Visibility.Collapsed))

        Public Property VisibilityCheckSyntaxBlock() As Visibility
            Get
                Return DirectCast(GetValue(VisibilityCheckSyntaxBlockProperty), Visibility)
            End Get
            Set(value As Visibility)
                SetValue(VisibilityCheckSyntaxBlockProperty, value)
            End Set
        End Property

        Public Sub New()
            InitializeComponent()

            Visibility = Visibility.Collapsed

            Dim collection = New ObservableCollection(Of ComboBoxItem)()
            For Each syntax As Type In ActiveQueryBuilder.Core.Helpers.SyntaxProviderList
                Dim instance = TryCast(Activator.CreateInstance(syntax), BaseSyntaxProvider)
                collection.Add(New ComboBoxItem(instance))
            Next syntax

            ComboBoxSyntaxProvider.ItemsSource = collection
        End Sub

        Public Sub Show(message As String, currentSyntaxProvider As BaseSyntaxProvider)
            If String.IsNullOrEmpty(message) Then
                Visibility = Visibility.Collapsed
                Return
            End If

            _allowChangedSyntax = False
            TextBlockErrorPrompt.Text = message
            ComboBoxSyntaxProvider.Text = currentSyntaxProvider.ToString()
            _allowChangedSyntax = True
            Visibility = Visibility.Visible
        End Sub

        Private Sub ComboBoxSyntaxProvider_OnSelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            If Not _allowChangedSyntax Then
                Return
            End If

            Dim syntaxProvider = CType(ComboBoxSyntaxProvider.SelectedItem, ComboBoxItem).SyntaxProvider

            OnSyntaxProviderChanged(New SelectionChangedEventArgs(e.RoutedEvent, New List(Of BaseSyntaxProvider)(), New List(Of BaseSyntaxProvider) From {syntaxProvider}))
        End Sub

        Protected Overridable Sub OnSyntaxProviderChanged(e As SelectionChangedEventArgs)
            SyntaxProviderChangedEvent?.Invoke(Me, e)
            Visibility = Visibility.Collapsed
        End Sub

        Private Sub HyperlinkGoToPosition_OnClick(sender As Object, e As RoutedEventArgs)
            OnGoToErrorPositionEvent()
            Visibility = Visibility.Collapsed
        End Sub

        Private Sub HyperlinkPreviousValidText_OnClick(sender As Object, e As RoutedEventArgs)
            OnRevertValidTextEvent()
            Visibility = Visibility.Collapsed
        End Sub

        Protected Overridable Sub OnGoToErrorPositionEvent()
            GoToErrorPositionEvent?.Invoke(Me, EventArgs.Empty)
        End Sub

        Protected Overridable Sub OnRevertValidTextEvent()
            RevertValidTextEvent?.Invoke(Me, EventArgs.Empty)
        End Sub
    End Class

    Public Class ComboBoxItem
        Public ReadOnly Property SyntaxProvider() As BaseSyntaxProvider
        Public ReadOnly Property DisplayString() As String
            Get
                Return SyntaxProvider.ToString()
            End Get
        End Property
        Public Sub New()
        End Sub

        Public Sub New(provider As BaseSyntaxProvider)
            SyntaxProvider = provider
        End Sub
    End Class
End Namespace
