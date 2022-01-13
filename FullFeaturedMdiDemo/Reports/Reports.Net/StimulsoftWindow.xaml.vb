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
Imports System.Data
Imports System.Drawing
Imports System.Windows
Imports ActiveQueryBuilder.View.WPF.Annotations
Imports Stimulsoft.Base.Drawing
Imports Stimulsoft.Report
Imports Stimulsoft.Report.Components

Partial Public Class StimulsoftWindow
    Private Property DataTable As DataTable

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(
<NotNull> ByVal dataTable As DataTable)
        Me.New()
        dataTable = dataTable
    End Sub

    Private Sub ShowDesigner_OnClick(ByVal sender As Object, ByVal e As RoutedEventArgs)
        ReportViewer.Report.Design()
    End Sub

    Private Sub StimulsoftWindow_OnLoaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim report As StiReport = New StiReport()
        report.RegData(DataTable)
        report.Dictionary.Synchronize()
        Dim page = report.Pages(0)
        Dim headerBand = New StiHeaderBand With {
            .Height = 0.5,
            .Name = "HeaderBand"
        }
        page.Components.Add(headerBand)
        Dim dataBand = New StiDataBand With {
            .DataSourceName = "result",
            .Height = 0.5,
            .Name = "DataBand"
        }
        page.Components.Add(dataBand)
        Dim width = page.Width / DataTable.Columns.Count

        For Each column As DataColumn In DataTable.Columns
            Dim headerText = New StiText(New RectangleD(0, 0, width, 0.5)) With {
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
            Dim dataText = New StiText(New RectangleD(0, 0, width, 0.5)) With {
                .Text = "{result." & column.ColumnName & "}",
                .Dockable = True,
                .DockStyle = StiDockStyle.Left,
                .VertAlignment = StiVertAlignment.Center,
                .CanShrink = True,
                .CanGrow = True
            }
            dataBand.Components.Add(dataText)
        Next

        ReportViewer.Report = report
        report.Compile()
        report.Render(True)
    End Sub
End Class
