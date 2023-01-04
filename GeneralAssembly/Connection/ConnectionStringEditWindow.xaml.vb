''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2023 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Namespace Connection
    ''' <summary>
    ''' Interaction logic for ConnectionStringEditWindow.xaml
    ''' </summary>
    Partial Public Class ConnectionStringEditWindow
        Public Property ConnectionString() As String
            Get
                Return tbConnectionString.Text
            End Get
            Set(value As String)
                tbConnectionString.Text = value
                Modified = False
            End Set
        End Property

        Private privateModified As Boolean
        Public Property Modified() As Boolean
            Get
                Return privateModified
            End Get
            Private Set(value As Boolean)
                privateModified = value
            End Set
        End Property

        Public Sub New()
            InitializeComponent()
            Modified = True
        End Sub

        Private Sub ButtonBaseOk_OnClick(sender As Object, e As RoutedEventArgs)
            DialogResult = True
            Close()
        End Sub

        Private Sub ButtonBaseClose_OnClick(sender As Object, e As RoutedEventArgs)
            Close()
        End Sub
    End Class
End Namespace
