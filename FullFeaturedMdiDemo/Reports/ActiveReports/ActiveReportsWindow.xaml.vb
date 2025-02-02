''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2025 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports ActiveQueryBuilder.View.WPF.Annotations
Imports GrapeCity.ActiveReports
Imports GrapeCity.ActiveReports.Document.Section
Imports GrapeCity.ActiveReports.SectionReportModel

Partial Public Class ActiveReportsWindow
    Private Property DataTable As DataTable
    Public Sub New(<NotNull> dataTable As DataTable)

        Me.DataTable = dataTable
        InitializeComponent()
    End Sub

    Private _sectionReport As SectionReport

    Private Sub ActiveReportsWindow_OnLoaded(sender As Object, e As RoutedEventArgs)
        _sectionReport = New SectionReport With {.DataSource = DataTable}

        _sectionReport.Document.Printer.PrinterName = ""
        _sectionReport.PageSettings.Margins.Left = 0.1F
        _sectionReport.PageSettings.Margins.Right = 0.1F
        _sectionReport.PrintWidth = 15.0F

        _sectionReport.Sections.Add(SectionType.ReportHeader, "ReportHeader")
        _sectionReport.Sections.Add(SectionType.Detail, "Detail")
        _sectionReport.Sections.Add(SectionType.ReportFooter, "ReportFooter")

        Dim section1 = _sectionReport.Sections(1)
        section1.CanGrow = True
        section1.CanShrink = True

        Dim locationX As Single = 0

        For Each dataSetColumn As DataColumn In DataTable.Columns
            Dim labelHeader = New Label With {
                .Text = dataSetColumn.ColumnName,
                .Alignment = GrapeCity.ActiveReports.Document.Section.TextAlignment.Left,
                .Location = New PointF(locationX, 0.0F),
                .ShrinkToFit = False,
                .MinCondenseRate = 100,
                .BackColor = System.Drawing.Color.Gainsboro,
                .Width = 2
            }
            _sectionReport.Sections(0).Controls.Add(labelHeader)

            Dim ctl = New TextBox With {
                .Name = dataSetColumn.ColumnName,
                .Text = dataSetColumn.ColumnName,
                .DataField = dataSetColumn.ColumnName,
                .Location = New PointF(locationX, 0.05F),
                .ShrinkToFit = False,
                .WrapMode = GrapeCity.ActiveReports.Document.Section.WrapMode.NoWrap,
                .CanShrink = False
            }

            ctl.Border.BottomStyle = BorderLineStyle.Dash

            _sectionReport.Sections(1).Controls.Add(ctl)

            locationX += ctl.Width
        Next dataSetColumn

        ReportViewer.LoadDocument(_sectionReport)
    End Sub
End Class
