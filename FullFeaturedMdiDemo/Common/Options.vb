'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports System.Windows
Imports System.Xml
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.Core.Serialization
Imports ActiveQueryBuilder.View
Imports ActiveQueryBuilder.View.WPF
Imports ActiveQueryBuilder.View.WPF.ExpressionEditor
Imports ActiveQueryBuilder.View.WPF.QueryView
Imports ActiveQueryBuilder.View.WPF.Serialization

Namespace Common
    Public Class Options
        Private ReadOnly DefaultTextEditorPadding As Thickness = New Thickness(5, 5, 0, 0)
        Public Property BehaviorOptions As BehaviorOptions
        Public Property DatabaseSchemaViewOptions As DatabaseSchemaViewOptions
        Public Property DesignPaneOptions As DesignPaneOptions
        Public Property VisualOptions As VisualOptions
        Public Property AddObjectDialogOptions As AddObjectDialogOptions
        Public Property DataSourceOptions As DataSourceOptions
        Public Property QueryColumnListOptions As QueryColumnListOptions
        Public Property QueryNavBarOptions As QueryNavBarOptions
        Public Property UserInterfaceOptions As UserInterfaceOptions
        Public Property SqlFormattingOptions As SQLFormattingOptions
        Public Property SqlGenerationOptions As SQLGenerationOptions
        Public Property ExpressionEditorOptions As ExpressionEditorOptions
        Public Property TextEditorOptions As TextEditorOptions
        Public Property TextEditorSqlOptions As SqlTextEditorOptions
        Private ReadOnly _options As List(Of OptionsBase) = New List(Of OptionsBase)()

        Public Sub New()
            CreateDefaultOptions()
        End Sub

        Public Sub CreateDefaultOptions()
            BehaviorOptions = New BehaviorOptions()
            DatabaseSchemaViewOptions = New DatabaseSchemaViewOptions()
            DesignPaneOptions = New DesignPaneOptions()
            VisualOptions = New VisualOptions()
            AddObjectDialogOptions = New AddObjectDialogOptions()
            DataSourceOptions = New DataSourceOptions()
            QueryColumnListOptions = New QueryColumnListOptions()
            QueryNavBarOptions = New QueryNavBarOptions()
            UserInterfaceOptions = New UserInterfaceOptions()
            SqlFormattingOptions = New SQLFormattingOptions()
            SqlGenerationOptions = New SQLGenerationOptions()
            ExpressionEditorOptions = New ExpressionEditorOptions()
            TextEditorOptions = New TextEditorOptions With {
                .Padding = DefaultTextEditorPadding,
                .LineHeight = New LengthUnit(90, SizeUnitType.Percent)
            }
            TextEditorSqlOptions = New SqlTextEditorOptions()
        End Sub

        Private Sub InitializeOptionsList()
            _options.Clear()
            _options.Add(BehaviorOptions)
            _options.Add(DatabaseSchemaViewOptions)
            _options.Add(DesignPaneOptions)
            _options.Add(VisualOptions)
            _options.Add(AddObjectDialogOptions)
            _options.Add(DataSourceOptions)
            _options.Add(QueryColumnListOptions)
            _options.Add(QueryNavBarOptions)
            _options.Add(UserInterfaceOptions)
            _options.Add(SqlFormattingOptions)
            _options.Add(SqlGenerationOptions)
            _options.Add(ExpressionEditorOptions)
            _options.Add(TextEditorOptions)
            _options.Add(TextEditorSqlOptions)
        End Sub

        Public Function SerializeToString() As String
            InitializeOptionsList()
            Dim result As String

            Using stream As MemoryStream = New MemoryStream()

                Using xmlBuilder As XmlDescriptionBuilder = New XmlDescriptionBuilder(stream)
                    Dim service As OptionsSerializationService = New OptionsSerializationService(xmlBuilder) With {
                        .SerializeDefaultValues = True
                    }
                    XmlSerializerExtensions.Builder = xmlBuilder
                    Dim root As Object = xmlBuilder.BeginObject("Options")

                    If True Then

                        For Each opt As OptionsBase In _options
                            Dim optionHandle As Object = xmlBuilder.BeginObjectProperty(root, opt.[GetType]().Name)

                            If True Then
                                service.EncodeObject(optionHandle, opt)
                            End If

                            xmlBuilder.EndObjectProperty(optionHandle)
                        Next
                    End If

                    xmlBuilder.EndObject(root)
                End Using

                stream.Position = 0

                Using reader As StreamReader = New StreamReader(stream)
                    result = reader.ReadToEnd()
                End Using
            End Using

            Return result
        End Function

        Public Sub SerializeToFile(path As String)
            InitializeOptionsList()
            File.WriteAllText(path, SerializeToString())
        End Sub

        Public Sub DeserializeFromFile(path As String)
            InitializeOptionsList()
            DeserializeFromString(File.ReadAllText(path))
        End Sub

        Public Sub DeserializeFromString(xml As String)
            InitializeOptionsList()
            Dim buffer As Byte() = Encoding.UTF8.GetBytes(xml)

            Using memoryStream As MemoryStream= New MemoryStream(buffer)
                Dim adapter As XmlAdapter = New XmlAdapter(memoryStream)
                Dim service As OptionsDeserializationService = New OptionsDeserializationService(adapter)
                XmlSerializerExtensions.Adapter = adapter
                adapter.Reader.ReadToFollowing(_options(0).[GetType]().Name)

                For Each opt As OptionsBase In _options
                    Dim optionTree As XmlReader = adapter.Reader.ReadSubtree()
                    optionTree.Read()
                    service.DecodeObject(optionTree, opt)
                    optionTree.Close()
                    adapter.Reader.Read()
                Next
            End Using
        End Sub
    End Class
End Namespace
