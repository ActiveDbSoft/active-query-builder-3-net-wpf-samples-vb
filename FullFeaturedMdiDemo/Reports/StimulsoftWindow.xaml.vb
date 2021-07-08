''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2021 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.Drawing
Imports ActiveQueryBuilder.View.WPF.Annotations
Imports Stimulsoft.Base.Drawing
Imports Stimulsoft.Report
Imports Stimulsoft.Report.Components

Namespace Reports
    Partial Public Class StimulsoftWindow

        Private Property DataTable() As DataTable

        Public Sub New(<NotNull> dataTable As DataTable)
            Me.New()
            Me.DataTable = dataTable
        End Sub

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub StimulsoftWindow_OnLoaded(sender As Object, e As RoutedEventArgs)
            Dim report As New StiReport()

            ' Add data to datastore
            report.RegData(DataTable)

            ' Fill dictionary
            report.Dictionary.Synchronize()

            Dim page = report.Pages(0)

            'Create HeaderBand
            Dim headerBand = New StiHeaderBand With {
                .Height = 0.5,
                .Name = "HeaderBand"
            }
            page.Components.Add(headerBand)

            'Create Databand
            Dim dataBand = New StiDataBand With {
                .DataSourceName = "result",
                .Height = 0.5,
                .Name = "DataBand"
            }
            page.Components.Add(dataBand)

            Dim width_Renamed = CInt(page.Width) \ DataTable.Columns.Count
            For Each column As DataColumn In DataTable.Columns
                'Create text on header
                Dim headerText = New StiText(New RectangleD(0, 0, width_Renamed, 0.5)) With {
                        .Text = column.ColumnName,
                        .HorAlignment = StiTextHorAlignment.Center,
                        .Brush = New StiSolidBrush(Color.Gainsboro),
                        .Dockable = True,
                        .DockStyle = StiDockStyle.Left,
                        .CanShrink = True,
                        .CanGrow = True,
                        .VertAlignment = StiVertAlignment.Center
                }
                headerBand.Components.Add(headerText)

                'Create text
                Dim dataText = New StiText(New RectangleD(0, 0, width_Renamed, 0.5)) With {
                        .Text = "{result." + column.ColumnName + "}",
                        .Dockable = True,
                        .DockStyle = StiDockStyle.Left,
                        .VertAlignment = StiVertAlignment.Center,
                        .CanShrink = True,
                        .CanGrow = True
                }

                dataBand.Components.Add(dataText)
            Next column

            Viewer.Report = report
            report.Compile()
            report.Render(True)
        End Sub

        Private Sub ShowDesigner_OnClick(sender As Object, e As RoutedEventArgs)
            Dim result = Viewer.Report.Design()
        End Sub
    End Class
End Namespace
