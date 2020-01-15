'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Windows
Imports System.Windows.Documents

Namespace Common
    Public Partial Class ExceptionWindow
        Public Property Message As String
            Get
                Return BlockMessage.Text
            End Get
            Set(value As String)
                BlockMessage.Text = value
            End Set
        End Property

        Public Property StackTrace As String
            Get
                Return New TextRange(BoxStackTrace.Document.ContentStart, BoxStackTrace.Document.ContentEnd).Text
            End Get
            Set(value As String)
                BoxStackTrace.Document.Blocks.Clear()
                BoxStackTrace.Document.Blocks.Add(New Paragraph(New Run(value)))
            End Set
        End Property

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub ButtonOk_OnClick(sender As Object, e As RoutedEventArgs)
            Close()
        End Sub
    End Class
End Namespace
