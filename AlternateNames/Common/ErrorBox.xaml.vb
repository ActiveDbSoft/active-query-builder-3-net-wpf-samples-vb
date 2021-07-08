''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2021 Active Database Software              ''
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

    Public Partial Class ErrorBox
        Private _allowChangedSyntax As Boolean = True
        Public Event SyntaxProviderChanged As SelectionChangedEventHandler
        Public Event GoToErrorPosition As EventHandler
        Public Event RevertValidText As EventHandler
        Public Shared ReadOnly VisibilityCheckSyntaxBlockProperty As DependencyProperty = DependencyProperty.Register("VisibilityCheckSyntaxBlock", GetType(Visibility), GetType(ErrorBox), New PropertyMetadata(Visibility.Collapsed))

        Public Property VisibilityCheckSyntaxBlock As Visibility
            Get
                Return CType(GetValue(VisibilityCheckSyntaxBlockProperty), Visibility)
            End Get
            Set(ByVal value As Visibility)
                SetValue(VisibilityCheckSyntaxBlockProperty, value)
            End Set
        End Property

        Public Sub New()
            InitializeComponent()
            Visibility = Visibility.Collapsed
            Dim collection = New ObservableCollection(Of ComboBoxItem)()

            For Each syntax As Type In Helpers.SyntaxProviderList
                Dim instance = TryCast(Activator.CreateInstance(syntax), BaseSyntaxProvider)
                collection.Add(New ComboBoxItem(instance))
            Next

            ComboBoxSyntaxProvider.ItemsSource = collection
        End Sub

        Public Sub Show(ByVal message As String, ByVal currentSyntaxProvider As BaseSyntaxProvider)
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

        Private Sub ComboBoxSyntaxProvider_OnSelectionChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
            If Not _allowChangedSyntax Then Return
            Dim syntaxProvider = (CType(ComboBoxSyntaxProvider.SelectedItem, ComboBoxItem)).SyntaxProvider
            OnSyntaxProviderChanged(New SelectionChangedEventArgs(e.RoutedEvent, New List(Of BaseSyntaxProvider)(), New List(Of BaseSyntaxProvider) From {
                                                                     syntaxProvider
                                                                     }))
        End Sub

        Protected Overridable Sub OnSyntaxProviderChanged(ByVal e As SelectionChangedEventArgs)
            RaiseEvent SyntaxProviderChanged(Me, e)
            Visibility = Visibility.Collapsed
        End Sub

        Private Sub HyperlinkGoToPosition_OnClick(ByVal sender As Object, ByVal e As RoutedEventArgs)
            OnGoToErrorPositionEvent()
            Visibility = Visibility.Collapsed
        End Sub

        Private Sub HyperlinkPreviousValidText_OnClick(ByVal sender As Object, ByVal e As RoutedEventArgs)
            OnRevertValidTextEvent()
            Visibility = Visibility.Collapsed
        End Sub

        Protected Overridable Sub OnGoToErrorPositionEvent()
            RaiseEvent GoToErrorPosition(Me, EventArgs.Empty)
        End Sub

        Protected Overridable Sub OnRevertValidTextEvent()
            RaiseEvent RevertValidText(Me, EventArgs.Empty)
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
End NameSpace
