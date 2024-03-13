''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2024 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Namespace Windows
    ''' <summary>
    ''' Interaction logic for AboutWindow.xaml
    ''' </summary>
    Partial Public Class AboutWindow
        Public Sub New()
            InitializeComponent()

            LblQueryBuilderVersion.Text = "v" & GetType(QueryBuilder).Assembly.GetName().Version.ToString()
            LblDemoVersion.Text = "v" & System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
        End Sub

        Private Sub Hyperlink1_RequestNavigate(sender As Object, e As RequestNavigateEventArgs)
            Process.Start(New ProcessStartInfo(e.Uri.AbsoluteUri))
            e.Handled = True
        End Sub
    End Class
End Namespace
