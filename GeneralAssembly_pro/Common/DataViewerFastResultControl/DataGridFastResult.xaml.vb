''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2025 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System
Imports System.Collections
Imports ActiveQueryBuilder.Core

Namespace Common.DataViewerFastResultControl
    Partial Public Class DataGridFastResult
        Public Property ItemsSource() As IEnumerable
            Get
                Return DGrid.ItemsSource
            End Get
            Set(value As IEnumerable)
                DGrid.ItemsSource = value
            End Set
        End Property

        Public Property ErrorMessage() As String
            Set(value As String)
                ErrorMessageBox.Message = value
            End Set
            Get
                Return ErrorMessageBox.Message
            End Get
        End Property
        Public Sub New()
            InitializeComponent()
        End Sub

        Public Sub FillData(resultSql As String, query As SQLQuery)
            Try
                ErrorMessage = ""

                Dim dv = SqlHelpers.GetDataView(resultSql, query)

                ItemsSource = dv
            Catch exception As Exception
                ErrorMessage = exception.Message
                ItemsSource = Nothing
            End Try
        End Sub
    End Class
End Namespace
