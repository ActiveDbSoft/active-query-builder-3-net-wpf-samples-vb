//*******************************************************************//
//       Active Query Builder Component Suite                        //
//                                                                   //
//       Copyright © 2006-2021 Active Database Software              //
//       ALL RIGHTS RESERVED                                         //
//                                                                   //
//       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            //
//       RESTRICTIONS.                                               //
//*******************************************************************//

Imports QueryStructureDemo.Info

''' <summary>
''' Interaction logic for MainWindow.xaml
''' </summary>
Partial Public Class MainWindow
    Private _lastValidSql As String
    Private _errorPosition As Integer = -1

    Private Const CSampleSelect As String = "Select 1 as cid, Upper('2'), 3, 4 + 1, 5 + 2 IntExpression "

    Private Const CSampleSelectFromWhere As String = "Select c.CustomerId as cid, c.CompanyName, Upper(c.CompanyName), o.OrderId " & "From Customers c Inner Join Orders o On c.CustomerID = o.CustomerID Where o.OrderId > 0 and c.CustomerId < 10"

    Private Const CSampleSelectFromJoin As String = "Select c.CustomerId as cid, Upper(c.CompanyName), o.OrderId, " & " p.ProductName + 1, 2+2 IntExpression From Customers c Inner Join " & "Orders o On c.CustomerID = o.CustomerID Inner Join " & "[Order Details] od On o.OrderID = od.OrderID Inner Join " & "Products p On p.ProductID = od.ProductID "

    Private Const CSampleSelectFromJoinSubqueries As String = "Select c.CustomerId as cid, Upper(c.CompanyName), o.OrderId, " & "p.ProductName + 1, 2+2 IntExpression From Customers c Inner Join " & "Orders o On c.CustomerID = o.CustomerID Inner Join " & "[Order Details] od On o.OrderID = od.OrderID Inner Join " & "(select pr.ProductId, pr.ProductName from Products pr) p On p.ProductID = od.ProductID "

    Private Const CSampleUnions As String = "Select c.CustomerId as cid, Upper(c.CompanyName), o.OrderId, " & "p.ProductName + 1, 2+2 IntExpression From Customers c Inner Join " & "Orders o On c.CustomerID = o.CustomerID Inner Join " & "[Order Details] od On o.OrderID = od.OrderID Inner Join " & "(select pr.ProductId, pr.ProductName from Products pr) p " & "On p.ProductID = od.ProductID union all " & "(select 1,2,3,4,5 union all select 6,7,8,9,0) union all " & "select (select Null as ""Null"") as EmptyValue, " & "SecondColumn = 2, Lower('ThirdColumn') as ThirdColumn, 0 as ""Quoted Alias"", 2+2*2 "


    Public Sub New()
        InitializeComponent()

        ' set required syntax provider
        Builder.SyntaxProvider = New MSSQLSyntaxProvider()

        AddHandler Loaded, AddressOf MainWindow_Loaded

    End Sub

    Private Sub MainWindow_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
        RemoveHandler Loaded, AddressOf MainWindow_Loaded

        ' Load sample database metadata from XML file
        Try
            Builder.MetadataLoadingOptions.OfflineMode = True
            Builder.MetadataContainer.ImportFromXML("Northwind.xml")
            Builder.InitializeDatabaseSchemaTree()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        ' load sample query
        Builder.SQL = CSampleSelectFromWhere

    End Sub

    Private Sub Builder_OnSQLUpdated(ByVal sender As Object, ByVal e As EventArgs)
        ' Builder generates new SQL query text. Show it to user.
        tbSQL.Text = Builder.FormattedSQL
        _lastValidSql = tbSQL.Text

        ' update info for entire query
        UpdateQueryInfo()
    End Sub

    Private Sub UpdateQueryInfo()
        ' update Query Structure information
        UpdateQueryStatisticsInfo()

        ' and update SubQueries list
        UpdateQuerySubQueriesInfo()

        ' and update information for current SubQuery/UnionSubQuery
        UpdateSubQueryInfo()
    End Sub

    Private Sub UpdateQueryStatisticsInfo()
        Dim statistics As QueryStatistics = Builder.QueryStatistics
        Dim stringBuilder As New StringBuilder()

        StatisticsInfo.DumpQueryStatisticsInfo(stringBuilder, statistics)

        tbStatistics.Text = stringBuilder.ToString()
    End Sub

    Private Sub UpdateQuerySubQueriesInfo()
        Dim stringBuilder = New StringBuilder()

        SubQueriesInfo.DumpSubQueriesInfo(stringBuilder, Builder)

        tbSubQueries.Text = stringBuilder.ToString()
    End Sub

    Private Sub UpdateSubQueryStructureInfo()
        Dim subQuery As SubQuery = Builder.ActiveUnionSubQuery.ParentSubQuery
        Dim stringBuilder As New StringBuilder()

        SubQueryStructureInfo.DumpQueryStructureInfo(stringBuilder, subQuery)

        tbSubQueryStructure.Text = stringBuilder.ToString()
    End Sub

    Private Sub UpdateWhereInfo()
        Dim unionSubQuery As UnionSubQuery = Builder.ActiveUnionSubQuery
        Dim stringBuilder As New StringBuilder()

        Dim unionSubQueryAst As SQLSubQuerySelectExpression = unionSubQuery.ResultQueryAST

        Try
            If unionSubQueryAst.Where IsNot Nothing Then
                WhereInfo.DumpWhereInfo(stringBuilder, unionSubQueryAst.Where)
            End If
        Finally
            unionSubQueryAst.Dispose()
        End Try

        tbWhere.Text = stringBuilder.ToString()
    End Sub

    Private Sub UpdateSelectedExpressionsInfo()
        Dim unionSubQuery As UnionSubQuery = Builder.ActiveUnionSubQuery
        Dim stringBuilder As New StringBuilder()

        SelectedExpressionsInfo.DumpSelectedExpressionsInfoFromUnionSubQuery(stringBuilder, unionSubQuery)

        tbSelectedExpressions.Text = stringBuilder.ToString()
    End Sub

    Private Sub UpdateSubQueryInfo()
        ' update Query Structure information
        UpdateSubQueryStructureInfo()

        ' update DataSources information
        UpdateDataSourcesInfo()

        ' update Links information
        UpdateLinksInfo()

        ' update Selected Expressions information
        UpdateSelectedExpressionsInfo()

        ' and update WHERE clause information
        UpdateWhereInfo()
    End Sub

    Private Sub UpdateDataSourcesInfo()
        Dim unionSubQuery As UnionSubQuery = Builder.ActiveUnionSubQuery
        Dim stringBuilder As New StringBuilder()

        DataSourcesInfo.DumpDataSourcesInfoFromUnionSubQuery(stringBuilder, unionSubQuery)

        tbDataSources.Text = stringBuilder.ToString()
    End Sub

    Private Sub UpdateLinksInfo()
        Dim unionSubQuery As UnionSubQuery = Builder.ActiveUnionSubQuery
        Dim stringBuilder As New StringBuilder()

        LinksInfo.DumpLinksInfoFromUnionSubQuery(stringBuilder, unionSubQuery)

        tbLinks.Text = stringBuilder.ToString()
    End Sub

    Private Sub Builder_OnActiveUnionSubQueryChanged(ByVal sender As Object, ByVal e As EventArgs)
        If Builder.ActiveUnionSubQuery Is Nothing Then
            Return
        End If
        ' update Query Structure information
        UpdateSubQueryStructureInfo()

        ' update DataSources information
        UpdateDataSourcesInfo()

        ' update Links information
        UpdateLinksInfo()

        ' and update Selected expressions information
        UpdateSelectedExpressionsInfo()
    End Sub

    Private Sub SqlTextBox_OnLostKeyboardFocus(ByVal sender As Object, ByVal e As KeyboardFocusChangedEventArgs)
        Try
            ' Update the query builder with manually edited query text:
            Builder.SQL = tbSQL.Text
            ErrorBox.Show(Nothing, Builder.SyntaxProvider)
        Catch ex As SQLParsingException
            ' Show banner with error text
            ErrorBox.Show(ex.Message, Builder.SyntaxProvider)

            _errorPosition = ex.ErrorPos.pos
        End Try
    End Sub

#Region "Button click"

    Private Sub btnSampleSelect_Click(ByVal sender As Object, ByVal e As EventArgs)
        Builder.SQL = CSampleSelect
    End Sub

    Private Sub btnSampleSelectFromWhere_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Builder.SQL = CSampleSelectFromWhere
    End Sub

    Private Sub btnSampleSelectFromJoin_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Builder.SQL = CSampleSelectFromJoin
    End Sub

    Private Sub btnSampleSelectFromJoinSubqueries_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Builder.SQL = CSampleSelectFromJoinSubqueries
    End Sub

    Private Sub btnSampleUnions_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Builder.SQL = CSampleUnions
    End Sub

    Private Sub btnShowUnlinkedDatasourcesButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        ' get active UnionSubQuery
        Dim unionSubQuery = Builder.ActiveUnionSubQuery.ParentUnionSubQuery

        ' analize links and datasources
        Dim unlinkedDatasourcesInfo = UnlinkedDatasources.GetUnlinkedDataSourcesFromUnionSubQuery(unionSubQuery)

        MessageBox.Show(Me, unlinkedDatasourcesInfo)
    End Sub

#End Region

    Private Sub TbSQL_OnTextChanged(ByVal sender As Object, ByVal e As TextChangedEventArgs)
        ErrorBox.Show(Nothing, Builder.SyntaxProvider)
    End Sub

    Private Sub ErrorBox_OnGoToErrorPosition(ByVal sender As Object, ByVal e As EventArgs)
        tbSQL.Focus()

        If _errorPosition = -1 Then
            Return
        End If

        If tbSQL.LineCount <> 1 Then
            tbSQL.ScrollToLine(tbSQL.GetLineIndexFromCharacterIndex(_errorPosition))
        End If
        tbSQL.CaretIndex = _errorPosition
    End Sub

    Private Sub ErrorBox_OnRevertValidText(ByVal sender As Object, ByVal e As EventArgs)
        tbSQL.Text = _lastValidSql
        tbSQL.Focus()
    End Sub
End Class
