''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2022 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Windows
Imports System.Windows.Input

Namespace Common
    ''' <summary>
    ''' Interaction logic for WindowNameQuery.xaml
    ''' </summary>
    Partial Public Class WindowNameQuery
        Public ReadOnly Property NameQuery() As String
            Get
                Return TextBoxNameQuery.Text
            End Get
        End Property
        Public Sub New()
            InitializeComponent()

            AddHandler Loaded, AddressOf WindowNameQuery_Loaded
        End Sub

        Private Sub WindowNameQuery_Loaded(sender As Object, e As RoutedEventArgs)
            RemoveHandler Loaded, AddressOf WindowNameQuery_Loaded

            Keyboard.Focus(TextBoxNameQuery)
            TextBoxNameQuery.SelectionStart = 0
            TextBoxNameQuery.SelectionLength = TextBoxNameQuery.Text.Length
        End Sub

        Private Sub ButtonOk_OnClick(sender As Object, e As RoutedEventArgs)
            DialogResult = True
        End Sub

        Private Sub ButtonCancel_OnClick(sender As Object, e As RoutedEventArgs)
            DialogResult = False
        End Sub
    End Class
End Namespace
