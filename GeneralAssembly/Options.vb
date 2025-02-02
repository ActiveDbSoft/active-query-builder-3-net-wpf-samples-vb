''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2025 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports System.IO
Imports System.Text
Imports System.Windows
Imports ActiveQueryBuilder.Core
Imports ActiveQueryBuilder.Core.Serialization
Imports ActiveQueryBuilder.View
Imports ActiveQueryBuilder.View.WPF
Imports ActiveQueryBuilder.View.WPF.QueryView
Imports ActiveQueryBuilder.View.WPF.Serialization

Public Class Options
    Private ReadOnly DefaultTextEditorPadding As New Thickness(5, 5, 0, 0)

    Public Property BehaviorOptions() As BehaviorOptions
    Public Property DatabaseSchemaViewOptions() As DatabaseSchemaViewOptions
    Public Property DesignPaneOptions() As DesignPaneOptions
    Public Property VisualOptions() As VisualOptions
    Public Property AddObjectDialogOptions() As AddObjectDialogOptions
    Public Property DataSourceOptions() As DataSourceOptions
    Public Property QueryNavBarOptions() As QueryNavBarOptions
    Public Property UserInterfaceOptions() As UserInterfaceOptions
    Public Property SqlFormattingOptions() As SQLFormattingOptions
    Public Property SqlGenerationOptions() As SQLGenerationOptions

    Private ReadOnly _options As New List(Of OptionsBase)()

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
        QueryNavBarOptions = New QueryNavBarOptions()
        UserInterfaceOptions = New UserInterfaceOptions()
        SqlFormattingOptions = New SQLFormattingOptions()
        SqlGenerationOptions = New SQLGenerationOptions()
    End Sub

    Private Sub InitializeOptionsList()
        _options.Clear()
        _options.Add(BehaviorOptions)
        _options.Add(DatabaseSchemaViewOptions)
        _options.Add(DesignPaneOptions)
        _options.Add(VisualOptions)
        _options.Add(AddObjectDialogOptions)
        _options.Add(DataSourceOptions)
        _options.Add(QueryNavBarOptions)
        _options.Add(UserInterfaceOptions)
        _options.Add(SqlFormattingOptions)
        _options.Add(SqlGenerationOptions)
    End Sub

    Public Function SerializeToString() As String
        InitializeOptionsList()

        Dim result = String.Empty
        Using stream = New MemoryStream()
            Using xmlBuilder = New XmlDescriptionBuilder(stream)
                Dim service = New OptionsSerializationService(xmlBuilder) With {.SerializeDefaultValues = True}
                XmlSerializerExtensions.Builder = xmlBuilder
                Using root = xmlBuilder.BeginObject("Options")
                    For Each [option] In _options
                        Using optionHandle = xmlBuilder.BeginObjectProperty(root, [option].GetType().Name)
                            service.EncodeObject(optionHandle, [option])
                        End Using
                    Next [option]
                End Using
            End Using

            stream.Position = 0
            Using reader = New StreamReader(stream)
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

        Dim buffer = Encoding.UTF8.GetBytes(xml)
        Using memoryStream = New MemoryStream(buffer)
            Dim adapter = New XmlAdapter(memoryStream)
            Dim service = New OptionsDeserializationService(adapter)
            XmlSerializerExtensions.Adapter = adapter

            adapter.Reader.ReadToFollowing(_options(0).GetType().Name)

            For Each [option] In _options
                Dim optionTree = adapter.Reader.ReadSubtree()
                optionTree.Read()
                service.DecodeObject(optionTree, [option])
                optionTree.Close()
                adapter.Reader.Read()
            Next [option]
        End Using
    End Sub
End Class
