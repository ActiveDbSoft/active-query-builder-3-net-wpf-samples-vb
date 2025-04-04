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

#Disable Warning BC40008
' ReSharper disable UnusedMember.Global
' ReSharper disable MemberCanBePrivate.Global
' ReSharper disable UnusedAutoPropertyAccessor.Global
' ReSharper disable IntroduceOptionalParameters.Global
' ReSharper disable MemberCanBeProtected.Global
' ReSharper disable InconsistentNaming
' ReSharper disable CheckNamespace

Namespace Annotations
  ''' <summary>
  ''' Indicates that the value of the marked element could be <c>null</c> sometimes,
  ''' so the check for <c>null</c> is necessary before its usage.
  ''' </summary>
  ''' <example><code>
  ''' [CanBeNull] public object Test() { return null; }
  ''' public void UseTest() {
  '''   var p = Test();
  '''   var s = p.ToString(); // Warning: Possible 'System.NullReferenceException'
  ''' }
  ''' </code></example>
  <AttributeUsage(AttributeTargets.Method Or AttributeTargets.Parameter Or AttributeTargets.Property Or AttributeTargets.Delegate Or AttributeTargets.Field Or AttributeTargets.Event)>
  Public NotInheritable Class CanBeNullAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' Indicates that the value of the marked element could never be <c>null</c>.
  ''' </summary>
  ''' <example><code>
  ''' [NotNull] public object Foo() {
  '''   return null; // Warning: Possible 'null' assignment
  ''' }
  ''' </code></example>
  <AttributeUsage(AttributeTargets.Method Or AttributeTargets.Parameter Or AttributeTargets.Property Or AttributeTargets.Delegate Or AttributeTargets.Field Or AttributeTargets.Event)>
  Public NotInheritable Class NotNullAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' Can be appplied to symbols of types derived from IEnumerable as well as to symbols of Task
  ''' and Lazy classes to indicate that the value of a collection item, of the Task.Result property
  ''' or of the Lazy.Value property can never be null.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Method Or AttributeTargets.Parameter Or AttributeTargets.Property Or AttributeTargets.Delegate Or AttributeTargets.Field)>
  Public NotInheritable Class ItemNotNullAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' Can be appplied to symbols of types derived from IEnumerable as well as to symbols of Task
  ''' and Lazy classes to indicate that the value of a collection item, of the Task.Result property
  ''' or of the Lazy.Value property can be null.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Method Or AttributeTargets.Parameter Or AttributeTargets.Property Or AttributeTargets.Delegate Or AttributeTargets.Field)>
  Public NotInheritable Class ItemCanBeNullAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' Indicates that the marked method builds string by format pattern and (optional) arguments.
  ''' Parameter, which contains format string, should be given in constructor. The format string
  ''' should be in <see cref="string.Format(IFormatProvider,string,object[])"/>-like form.
  ''' </summary>
  ''' <example><code>
  ''' [StringFormatMethod("message")]
  ''' public void ShowError(string message, params object[] args) { /* do something */ }
  ''' public void Foo() {
  '''   ShowError("Failed: {0}"); // Warning: Non-existing argument in format string
  ''' }
  ''' </code></example>
  <AttributeUsage(AttributeTargets.Constructor Or AttributeTargets.Method Or AttributeTargets.Property Or AttributeTargets.Delegate)>
  Public NotInheritable Class StringFormatMethodAttribute
      Inherits Attribute

    ''' <param name="formatParameterName">
    ''' Specifies which parameter of an annotated method should be treated as format-string
    ''' </param>
    Public Sub New(formatParameterName As String)
      Me.FormatParameterName = formatParameterName
    End Sub

    Private privateFormatParameterName As String
    Public Property FormatParameterName() As String
        Get
            Return privateFormatParameterName
        End Get
        Private Set(value As String)
            privateFormatParameterName = value
        End Set
    End Property
  End Class

  ''' <summary>
  ''' For a parameter that is expected to be one of the limited set of values.
  ''' Specify fields of which type should be used as values for this parameter.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Property Or AttributeTargets.Field)>
  Public NotInheritable Class ValueProviderAttribute
      Inherits Attribute

    Public Sub New(name As String)
      Me.Name = name
    End Sub

    Private privateName As String
    <NotNull>
    Public Property Name() As String
        Get
            Return privateName
        End Get
        Private Set(value As String)
            privateName = value
        End Set
    End Property
  End Class

  ''' <summary>
  ''' Indicates that the function argument should be string literal and match one
  ''' of the parameters of the caller function. For example, ReSharper annotates
  ''' the parameter of <see cref="System.ArgumentNullException"/>.
  ''' </summary>
  ''' <example><code>
  ''' public void Foo(string param) {
  '''   if (param == null)
  '''     throw new ArgumentNullException("par"); // Warning: Cannot resolve symbol
  ''' }
  ''' </code></example>
  <AttributeUsage(AttributeTargets.Parameter)>
  Public NotInheritable Class InvokerParameterNameAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' Indicates that the method is contained in a type that implements
  ''' <c>System.ComponentModel.INotifyPropertyChanged</c> interface and this method
  ''' is used to notify that some property value changed.
  ''' </summary>
  ''' <remarks>
  ''' The method should be non-static and conform to one of the supported signatures:
  ''' <list>
  ''' <item><c>NotifyChanged(string)</c></item>
  ''' <item><c>NotifyChanged(params string[])</c></item>
  ''' <item><c>NotifyChanged{T}(Expression{Func{T}})</c></item>
  ''' <item><c>NotifyChanged{T,U}(Expression{Func{T,U}})</c></item>
  ''' <item><c>SetProperty{T}(ref T, T, string)</c></item>
  ''' </list>
  ''' </remarks>
  ''' <example><code>
  ''' public class Foo : INotifyPropertyChanged {
  '''   public event PropertyChangedEventHandler PropertyChanged;
  '''   [NotifyPropertyChangedInvocator]
  '''   protected virtual void NotifyChanged(string propertyName) { ... }
  '''
  '''   private string _name;
  '''   public string Name {
  '''     get { return _name; }
  '''     set { _name = value; NotifyChanged("LastName"); /* Warning */ }
  '''   }
  ''' }
  ''' </code>
  ''' Examples of generated notifications:
  ''' <list>
  ''' <item><c>NotifyChanged("Property")</c></item>
  ''' <item><c>NotifyChanged(() =&gt; Property)</c></item>
  ''' <item><c>NotifyChanged((VM x) =&gt; x.Property)</c></item>
  ''' <item><c>SetProperty(ref myField, value, "Property")</c></item>
  ''' </list>
  ''' </example>
  <AttributeUsage(AttributeTargets.Method)>
  Public NotInheritable Class NotifyPropertyChangedInvocatorAttribute
      Inherits Attribute

    Public Sub New()
    End Sub
    Public Sub New(parameterName As String)
      Me.ParameterName = parameterName
    End Sub

    Private privateParameterName As String
    Public Property ParameterName() As String
        Get
            Return privateParameterName
        End Get
        Private Set(value As String)
            privateParameterName = value
        End Set
    End Property
  End Class

  ''' <summary>
  ''' Describes dependency between method input and output.
  ''' </summary>
  ''' <syntax>
  ''' <p>Function Definition Table syntax:</p>
  ''' <list>
  ''' <item>FDT      ::= FDTRow [;FDTRow]*</item>
  ''' <item>FDTRow   ::= Input =&gt; Output | Output &lt;= Input</item>
  ''' <item>Input    ::= ParameterName: Value [, Input]*</item>
  ''' <item>Output   ::= [ParameterName: Value]* {halt|stop|void|nothing|Value}</item>
  ''' <item>Value    ::= true | false | null | notnull | canbenull</item>
  ''' </list>
  ''' If method has single input parameter, it's name could be omitted.<br/>
  ''' Using <c>halt</c> (or <c>void</c>/<c>nothing</c>, which is the same)
  ''' for method output means that the methos doesn't return normally.<br/>
  ''' <c>canbenull</c> annotation is only applicable for output parameters.<br/>
  ''' You can use multiple <c>[ContractAnnotation]</c> for each FDT row,
  ''' or use single attribute with rows separated by semicolon.<br/>
  ''' </syntax>
  ''' <examples><list>
  ''' <item><code>
  ''' [ContractAnnotation("=> halt")]
  ''' public void TerminationMethod()
  ''' </code></item>
  ''' <item><code>
  ''' [ContractAnnotation("halt &lt;= condition: false")]
  ''' public void Assert(bool condition, string text) // regular assertion method
  ''' </code></item>
  ''' <item><code>
  ''' [ContractAnnotation("s:null => true")]
  ''' public bool IsNullOrEmpty(string s) // string.IsNullOrEmpty()
  ''' </code></item>
  ''' <item><code>
  ''' // A method that returns null if the parameter is null,
  ''' // and not null if the parameter is not null
  ''' [ContractAnnotation("null => null; notnull => notnull")]
  ''' public object Transform(object data) 
  ''' </code></item>
  ''' <item><code>
  ''' [ContractAnnotation("s:null=>false; =>true,result:notnull; =>false, result:null")]
  ''' public bool TryParse(string s, out Person result)
  ''' </code></item>
  ''' </list></examples>
  <AttributeUsage(AttributeTargets.Method, AllowMultiple := True)>
  Public NotInheritable Class ContractAnnotationAttribute
      Inherits Attribute

    Public Sub New(<NotNull> contract As String)
        Me.New(contract, False)
    End Sub

    Public Sub New(<NotNull> contract As String, forceFullStates As Boolean)
      Me.Contract = contract
      Me.ForceFullStates = forceFullStates
    End Sub

    Private privateContract As String
    Public Property Contract() As String
        Get
            Return privateContract
        End Get
        Private Set(value As String)
            privateContract = value
        End Set
    End Property
    Private privateForceFullStates As Boolean
    Public Property ForceFullStates() As Boolean
        Get
            Return privateForceFullStates
        End Get
        Private Set(value As Boolean)
            privateForceFullStates = value
        End Set
    End Property
  End Class

  ''' <summary>
  ''' Indicates that marked element should be localized or not.
  ''' </summary>
  ''' <example><code>
  ''' [LocalizationRequiredAttribute(true)]
  ''' public class Foo {
  '''   private string str = "my string"; // Warning: Localizable string
  ''' }
  ''' </code></example>
  <AttributeUsage(AttributeTargets.All)>
  Public NotInheritable Class LocalizationRequiredAttribute
      Inherits Attribute

    Public Sub New()
        Me.New(True)
    End Sub
    Public Sub New(required As Boolean)
      Me.Required = required
    End Sub

    Private privateRequired As Boolean
    Public Property Required() As Boolean
        Get
            Return privateRequired
        End Get
        Private Set(value As Boolean)
            privateRequired = value
        End Set
    End Property
  End Class

  ''' <summary>
  ''' Indicates that the value of the marked type (or its derivatives)
  ''' cannot be compared using '==' or '!=' operators and <c>Equals()</c>
  ''' should be used instead. However, using '==' or '!=' for comparison
  ''' with <c>null</c> is always permitted.
  ''' </summary>
  ''' <example><code>
  ''' [CannotApplyEqualityOperator]
  ''' class NoEquality { }
  ''' class UsesNoEquality {
  '''   public void Test() {
  '''     var ca1 = new NoEquality();
  '''     var ca2 = new NoEquality();
  '''     if (ca1 != null) { // OK
  '''       bool condition = ca1 == ca2; // Warning
  '''     }
  '''   }
  ''' }
  ''' </code></example>
  <AttributeUsage(AttributeTargets.Interface Or AttributeTargets.Class Or AttributeTargets.Struct)>
  Public NotInheritable Class CannotApplyEqualityOperatorAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' When applied to a target attribute, specifies a requirement for any type marked
  ''' with the target attribute to implement or inherit specific type or types.
  ''' </summary>
  ''' <example><code>
  ''' [BaseTypeRequired(typeof(IComponent)] // Specify requirement
  ''' public class ComponentAttribute : Attribute { }
  ''' [Component] // ComponentAttribute requires implementing IComponent interface
  ''' public class MyComponent : IComponent { }
  ''' </code></example>
  <AttributeUsage(AttributeTargets.Class, AllowMultiple := True), BaseTypeRequired(GetType(Attribute))>
  Public NotInheritable Class BaseTypeRequiredAttribute
      Inherits Attribute

    Public Sub New(<NotNull> baseType As Type)
      Me.BaseType = baseType
    End Sub

    Private privateBaseType As Type
    <NotNull>
    Public Property BaseType() As Type
        Get
            Return privateBaseType
        End Get
        Private Set(value As Type)
            privateBaseType = value
        End Set
    End Property
  End Class

  ''' <summary>
  ''' Indicates that the marked symbol is used implicitly (e.g. via reflection, in external library),
  ''' so this symbol will not be marked as unused (as well as by other usage inspections).
  ''' </summary>
  <AttributeUsage(AttributeTargets.All)>
  Public NotInheritable Class UsedImplicitlyAttribute
      Inherits Attribute

    Public Sub New()
        Me.New(ImplicitUseKindFlags.Default, ImplicitUseTargetFlags.Default)
    End Sub

    Public Sub New(useKindFlags As ImplicitUseKindFlags)
        Me.New(useKindFlags, ImplicitUseTargetFlags.Default)
    End Sub

    Public Sub New(targetFlags As ImplicitUseTargetFlags)
        Me.New(ImplicitUseKindFlags.Default, targetFlags)
    End Sub

    Public Sub New(useKindFlags As ImplicitUseKindFlags, targetFlags As ImplicitUseTargetFlags)
      Me.UseKindFlags = useKindFlags
      Me.TargetFlags = targetFlags
    End Sub

    Private privateUseKindFlags As ImplicitUseKindFlags
    Public Property UseKindFlags() As ImplicitUseKindFlags
        Get
            Return privateUseKindFlags
        End Get
        Private Set(value As ImplicitUseKindFlags)
            privateUseKindFlags = value
        End Set
    End Property
    Private privateTargetFlags As ImplicitUseTargetFlags
    Public Property TargetFlags() As ImplicitUseTargetFlags
        Get
            Return privateTargetFlags
        End Get
        Private Set(value As ImplicitUseTargetFlags)
            privateTargetFlags = value
        End Set
    End Property
  End Class

  ''' <summary>
  ''' Should be used on attributes and causes ReSharper to not mark symbols marked with such attributes
  ''' as unused (as well as by other usage inspections)
  ''' </summary>
  <AttributeUsage(AttributeTargets.Class Or AttributeTargets.GenericParameter)>
  Public NotInheritable Class MeansImplicitUseAttribute
      Inherits Attribute

    Public Sub New()
        Me.New(ImplicitUseKindFlags.Default, ImplicitUseTargetFlags.Default)
    End Sub

    Public Sub New(useKindFlags As ImplicitUseKindFlags)
        Me.New(useKindFlags, ImplicitUseTargetFlags.Default)
    End Sub

    Public Sub New(targetFlags As ImplicitUseTargetFlags)
        Me.New(ImplicitUseKindFlags.Default, targetFlags)
    End Sub

    Public Sub New(useKindFlags As ImplicitUseKindFlags, targetFlags As ImplicitUseTargetFlags)
      Me.UseKindFlags = useKindFlags
      Me.TargetFlags = targetFlags
    End Sub

    Private privateUseKindFlags As ImplicitUseKindFlags
    <UsedImplicitly>
    Public Property UseKindFlags() As ImplicitUseKindFlags
        Get
            Return privateUseKindFlags
        End Get
        Private Set(value As ImplicitUseKindFlags)
            privateUseKindFlags = value
        End Set
    End Property
    Private privateTargetFlags As ImplicitUseTargetFlags
    <UsedImplicitly>
    Public Property TargetFlags() As ImplicitUseTargetFlags
        Get
            Return privateTargetFlags
        End Get
        Private Set(value As ImplicitUseTargetFlags)
            privateTargetFlags = value
        End Set
    End Property
  End Class

  <Flags>
  Public Enum ImplicitUseKindFlags
    [Default] = Access Or Assign Or InstantiatedWithFixedConstructorSignature
    ''' <summary>Only entity marked with attribute considered used.</summary>
    Access = 1
    ''' <summary>Indicates implicit assignment to a member.</summary>
    Assign = 2
    ''' <summary>
    ''' Indicates implicit instantiation of a type with fixed constructor signature.
    ''' That means any unused constructor parameters won't be reported as such.
    ''' </summary>
    InstantiatedWithFixedConstructorSignature = 4
    ''' <summary>Indicates implicit instantiation of a type.</summary>
    InstantiatedNoFixedConstructorSignature = 8
  End Enum

  ''' <summary>
  ''' Specify what is considered used implicitly when marked
  ''' with <see cref="MeansImplicitUseAttribute"/> or <see cref="UsedImplicitlyAttribute"/>.
  ''' </summary>
  <Flags>
  Public Enum ImplicitUseTargetFlags
    [Default] = Itself
    Itself = 1
    ''' <summary>Members of entity marked with attribute are considered used.</summary>
    Members = 2
    ''' <summary>Entity marked with attribute and all its members considered used.</summary>
    WithMembers = Itself Or Members
  End Enum

  ''' <summary>
  ''' This attribute is intended to mark publicly available API
  ''' which should not be removed and so is treated as used.
  ''' </summary>
  <MeansImplicitUse(ImplicitUseTargetFlags.WithMembers)>
  Public NotInheritable Class PublicAPIAttribute
      Inherits Attribute

    Public Sub New()
    End Sub
    Public Sub New(<NotNull> comment As String)
      Me.Comment = comment
    End Sub

    Private privateComment As String
    Public Property Comment() As String
        Get
            Return privateComment
        End Get
        Private Set(value As String)
            privateComment = value
        End Set
    End Property
  End Class

  ''' <summary>
  ''' Tells code analysis engine if the parameter is completely handled when the invoked method is on stack.
  ''' If the parameter is a delegate, indicates that delegate is executed while the method is executed.
  ''' If the parameter is an enumerable, indicates that it is enumerated while the method is executed.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter)>
  Public NotInheritable Class InstantHandleAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' Indicates that a method does not make any observable state changes.
  ''' The same as <c>System.Diagnostics.Contracts.PureAttribute</c>.
  ''' </summary>
  ''' <example><code>
  ''' [Pure] private int Multiply(int x, int y) { return x * y; }
  ''' public void Foo() {
  '''   const int a = 2, b = 2;
  '''   Multiply(a, b); // Waring: Return value of pure method is not used
  ''' }
  ''' </code></example>
  <AttributeUsage(AttributeTargets.Method)>
  Public NotInheritable Class PureAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' Indicates that a parameter is a path to a file or a folder within a web project.
  ''' Path can be relative or absolute, starting from web root (~).
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter)>
  Public NotInheritable Class PathReferenceAttribute
      Inherits Attribute

    Public Sub New()
    End Sub
    Public Sub New(<PathReference> basePath As String)
      Me.BasePath = basePath
    End Sub

    Private privateBasePath As String
    Public Property BasePath() As String
        Get
            Return privateBasePath
        End Get
        Private Set(value As String)
            privateBasePath = value
        End Set
    End Property
  End Class

  ''' <summary>
  ''' An extension method marked with this attribute is processed by ReSharper code completion
  ''' as a 'Source Template'. When extension method is completed over some expression, it's source code
  ''' is automatically expanded like a template at call site.
  ''' </summary>
  ''' <remarks>
  ''' Template method body can contain valid source code and/or special comments starting with '$'.
  ''' Text inside these comments is added as source code when the template is applied. Template parameters
  ''' can be used either as additional method parameters or as identifiers wrapped in two '$' signs.
  ''' Use the <see cref="MacroAttribute"/> attribute to specify macros for parameters.
  ''' </remarks>
  ''' <example>
  ''' In this example, the 'forEach' method is a source template available over all values
  ''' of enumerable types, producing ordinary C# 'foreach' statement and placing caret inside block:
  ''' <code>
  ''' [SourceTemplate]
  ''' public static void forEach&lt;T&gt;(this IEnumerable&lt;T&gt; xs) {
  '''   foreach (var x in xs) {
  '''      //$ $END$
  '''   }
  ''' }
  ''' </code>
  ''' </example>
  <AttributeUsage(AttributeTargets.Method)>
  Public NotInheritable Class SourceTemplateAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' Allows specifying a macro for a parameter of a <see cref="SourceTemplateAttribute">source template</see>.
  ''' </summary>
  ''' <remarks>
  ''' You can apply the attribute on the whole method or on any of its additional parameters. The macro expression
  ''' is defined in the <see cref="MacroAttribute.Expression"/> property. When applied on a method, the target
  ''' template parameter is defined in the <see cref="MacroAttribute.Target"/> property. To apply the macro silently
  ''' for the parameter, set the <see cref="MacroAttribute.Editable"/> property value = -1.
  ''' </remarks>
  ''' <example>
  ''' Applying the attribute on a source template method:
  ''' <code>
  ''' [SourceTemplate, Macro(Target = "item", Expression = "suggestVariableName()")]
  ''' public static void forEach&lt;T&gt;(this IEnumerable&lt;T&gt; collection) {
  '''   foreach (var item in collection) {
  '''     //$ $END$
  '''   }
  ''' }
  ''' </code>
  ''' Applying the attribute on a template method parameter:
  ''' <code>
  ''' [SourceTemplate]
  ''' public static void something(this Entity x, [Macro(Expression = "guid()", Editable = -1)] string newguid) {
  '''   /*$ var $x$Id = "$newguid$" + x.ToString();
  '''   x.DoSomething($x$Id); */
  ''' }
  ''' </code>
  ''' </example>
  <AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Method, AllowMultiple := True)>
  Public NotInheritable Class MacroAttribute
      Inherits Attribute

    ''' <summary>
    ''' Allows specifying a macro that will be executed for a <see cref="SourceTemplateAttribute">source template</see>
    ''' parameter when the template is expanded.
    ''' </summary>
    Public Property Expression() As String

    ''' <summary>
    ''' Allows specifying which occurrence of the target parameter becomes editable when the template is deployed.
    ''' </summary>
    ''' <remarks>
    ''' If the target parameter is used several times in the template, only one occurrence becomes editable;
    ''' other occurrences are changed synchronously. To specify the zero-based index of the editable occurrence,
    ''' use values >= 0. To make the parameter non-editable when the template is expanded, use -1.
    ''' </remarks>>
    Public Property Editable() As Integer

    ''' <summary>
    ''' Identifies the target parameter of a <see cref="SourceTemplateAttribute">source template</see> if the
    ''' <see cref="MacroAttribute"/> is applied on a template method.
    ''' </summary>
    Public Property Target() As String
  End Class

  <AttributeUsage(AttributeTargets.Assembly, AllowMultiple := True)>
  Public NotInheritable Class AspMvcAreaMasterLocationFormatAttribute
      Inherits Attribute

    Public Sub New(format As String)
      Me.Format = format
    End Sub

    Private privateFormat As String
    Public Property Format() As String
        Get
            Return privateFormat
        End Get
        Private Set(value As String)
            privateFormat = value
        End Set
    End Property
  End Class

  <AttributeUsage(AttributeTargets.Assembly, AllowMultiple := True)>
  Public NotInheritable Class AspMvcAreaPartialViewLocationFormatAttribute
      Inherits Attribute

    Public Sub New(format As String)
      Me.Format = format
    End Sub

    Private privateFormat As String
    Public Property Format() As String
        Get
            Return privateFormat
        End Get
        Private Set(value As String)
            privateFormat = value
        End Set
    End Property
  End Class

  <AttributeUsage(AttributeTargets.Assembly, AllowMultiple := True)>
  Public NotInheritable Class AspMvcAreaViewLocationFormatAttribute
      Inherits Attribute

    Public Sub New(format As String)
      Me.Format = format
    End Sub

    Private privateFormat As String
    Public Property Format() As String
        Get
            Return privateFormat
        End Get
        Private Set(value As String)
            privateFormat = value
        End Set
    End Property
  End Class

  <AttributeUsage(AttributeTargets.Assembly, AllowMultiple := True)>
  Public NotInheritable Class AspMvcMasterLocationFormatAttribute
      Inherits Attribute

    Public Sub New(format As String)
      Me.Format = format
    End Sub

    Private privateFormat As String
    Public Property Format() As String
        Get
            Return privateFormat
        End Get
        Private Set(value As String)
            privateFormat = value
        End Set
    End Property
  End Class

  <AttributeUsage(AttributeTargets.Assembly, AllowMultiple := True)>
  Public NotInheritable Class AspMvcPartialViewLocationFormatAttribute
      Inherits Attribute

    Public Sub New(format As String)
      Me.Format = format
    End Sub

    Private privateFormat As String
    Public Property Format() As String
        Get
            Return privateFormat
        End Get
        Private Set(value As String)
            privateFormat = value
        End Set
    End Property
  End Class

  <AttributeUsage(AttributeTargets.Assembly, AllowMultiple := True)>
  Public NotInheritable Class AspMvcViewLocationFormatAttribute
      Inherits Attribute

    Public Sub New(format As String)
      Me.Format = format
    End Sub

    Private privateFormat As String
    Public Property Format() As String
        Get
            Return privateFormat
        End Get
        Private Set(value As String)
            privateFormat = value
        End Set
    End Property
  End Class

  ''' <summary>
  ''' ASP.NET MVC attribute. If applied to a parameter, indicates that the parameter
  ''' is an MVC action. If applied to a method, the MVC action name is calculated
  ''' implicitly from the context. Use this attribute for custom wrappers similar to
  ''' <c>System.Web.Mvc.Html.ChildActionExtensions.RenderAction(HtmlHelper, String)</c>.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Method)>
  Public NotInheritable Class AspMvcActionAttribute
      Inherits Attribute

    Public Sub New()
    End Sub
    Public Sub New(anonymousProperty As String)
      Me.AnonymousProperty = anonymousProperty
    End Sub

    Private privateAnonymousProperty As String
    Public Property AnonymousProperty() As String
        Get
            Return privateAnonymousProperty
        End Get
        Private Set(value As String)
            privateAnonymousProperty = value
        End Set
    End Property
  End Class

  ''' <summary>
  ''' ASP.NET MVC attribute. Indicates that a parameter is an MVC area.
  ''' Use this attribute for custom wrappers similar to
  ''' <c>System.Web.Mvc.Html.ChildActionExtensions.RenderAction(HtmlHelper, String)</c>.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter)>
  Public NotInheritable Class AspMvcAreaAttribute
      Inherits Attribute

    Public Sub New()
    End Sub
    Public Sub New(anonymousProperty As String)
      Me.AnonymousProperty = anonymousProperty
    End Sub

    Private privateAnonymousProperty As String
    Public Property AnonymousProperty() As String
        Get
            Return privateAnonymousProperty
        End Get
        Private Set(value As String)
            privateAnonymousProperty = value
        End Set
    End Property
  End Class

  ''' <summary>
  ''' ASP.NET MVC attribute. If applied to a parameter, indicates that the parameter is
  ''' an MVC controller. If applied to a method, the MVC controller name is calculated
  ''' implicitly from the context. Use this attribute for custom wrappers similar to
  ''' <c>System.Web.Mvc.Html.ChildActionExtensions.RenderAction(HtmlHelper, String, String)</c>.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Method)>
  Public NotInheritable Class AspMvcControllerAttribute
      Inherits Attribute

    Public Sub New()
    End Sub
    Public Sub New(anonymousProperty As String)
      Me.AnonymousProperty = anonymousProperty
    End Sub

    Private privateAnonymousProperty As String
    Public Property AnonymousProperty() As String
        Get
            Return privateAnonymousProperty
        End Get
        Private Set(value As String)
            privateAnonymousProperty = value
        End Set
    End Property
  End Class

  ''' <summary>
  ''' ASP.NET MVC attribute. Indicates that a parameter is an MVC Master. Use this attribute
  ''' for custom wrappers similar to <c>System.Web.Mvc.Controller.View(String, String)</c>.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter)>
  Public NotInheritable Class AspMvcMasterAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' ASP.NET MVC attribute. Indicates that a parameter is an MVC model type. Use this attribute
  ''' for custom wrappers similar to <c>System.Web.Mvc.Controller.View(String, Object)</c>.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter)>
  Public NotInheritable Class AspMvcModelTypeAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' ASP.NET MVC attribute. If applied to a parameter, indicates that the parameter is an MVC
  ''' partial view. If applied to a method, the MVC partial view name is calculated implicitly
  ''' from the context. Use this attribute for custom wrappers similar to
  ''' <c>System.Web.Mvc.Html.RenderPartialExtensions.RenderPartial(HtmlHelper, String)</c>.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Method)>
  Public NotInheritable Class AspMvcPartialViewAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' ASP.NET MVC attribute. Allows disabling inspections for MVC views within a class or a method.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Class Or AttributeTargets.Method)>
  Public NotInheritable Class AspMvcSupressViewErrorAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' ASP.NET MVC attribute. Indicates that a parameter is an MVC display template.
  ''' Use this attribute for custom wrappers similar to 
  ''' <c>System.Web.Mvc.Html.DisplayExtensions.DisplayForModel(HtmlHelper, String)</c>.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter)>
  Public NotInheritable Class AspMvcDisplayTemplateAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' ASP.NET MVC attribute. Indicates that a parameter is an MVC editor template.
  ''' Use this attribute for custom wrappers similar to
  ''' <c>System.Web.Mvc.Html.EditorExtensions.EditorForModel(HtmlHelper, String)</c>.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter)>
  Public NotInheritable Class AspMvcEditorTemplateAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' ASP.NET MVC attribute. Indicates that a parameter is an MVC template.
  ''' Use this attribute for custom wrappers similar to
  ''' <c>System.ComponentModel.DataAnnotations.UIHintAttribute(System.String)</c>.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter)>
  Public NotInheritable Class AspMvcTemplateAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' ASP.NET MVC attribute. If applied to a parameter, indicates that the parameter
  ''' is an MVC view. If applied to a method, the MVC view name is calculated implicitly
  ''' from the context. Use this attribute for custom wrappers similar to
  ''' <c>System.Web.Mvc.Controller.View(Object)</c>.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Method)>
  Public NotInheritable Class AspMvcViewAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' ASP.NET MVC attribute. When applied to a parameter of an attribute,
  ''' indicates that this parameter is an MVC action name.
  ''' </summary>
  ''' <example><code>
  ''' [ActionName("Foo")]
  ''' public ActionResult Login(string returnUrl) {
  '''   ViewBag.ReturnUrl = Url.Action("Foo"); // OK
  '''   return RedirectToAction("Bar"); // Error: Cannot resolve action
  ''' }
  ''' </code></example>
  <AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Property)>
  Public NotInheritable Class AspMvcActionSelectorAttribute
      Inherits Attribute

  End Class

  <AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Property Or AttributeTargets.Field)>
  Public NotInheritable Class HtmlElementAttributesAttribute
      Inherits Attribute

    Public Sub New()
    End Sub
    Public Sub New(name As String)
      Me.Name = name
    End Sub

    Private privateName As String
    Public Property Name() As String
        Get
            Return privateName
        End Get
        Private Set(value As String)
            privateName = value
        End Set
    End Property
  End Class

  <AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Field Or AttributeTargets.Property)>
  Public NotInheritable Class HtmlAttributeValueAttribute
      Inherits Attribute

    Public Sub New(<NotNull> name As String)
      Me.Name = name
    End Sub

    Private privateName As String
    <NotNull>
    Public Property Name() As String
        Get
            Return privateName
        End Get
        Private Set(value As String)
            privateName = value
        End Set
    End Property
  End Class

  ''' <summary>
  ''' Razor attribute. Indicates that a parameter or a method is a Razor section.
  ''' Use this attribute for custom wrappers similar to 
  ''' <c>System.Web.WebPages.WebPageBase.RenderSection(String)</c>.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Method)>
  Public NotInheritable Class RazorSectionAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' Indicates how method, constructor invocation or property access
  ''' over collection type affects content of the collection.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Method Or AttributeTargets.Constructor Or AttributeTargets.Property)>
  Public NotInheritable Class CollectionAccessAttribute
      Inherits Attribute

    Public Sub New(collectionAccessType As CollectionAccessType)
      Me.CollectionAccessType = collectionAccessType
    End Sub

    Private privateCollectionAccessType As CollectionAccessType
    Public Property CollectionAccessType() As CollectionAccessType
        Get
            Return privateCollectionAccessType
        End Get
        Private Set(value As CollectionAccessType)
            privateCollectionAccessType = value
        End Set
    End Property
  End Class

  <Flags>
  Public Enum CollectionAccessType
    ''' <summary>Method does not use or modify content of the collection.</summary>
    None = 0
    ''' <summary>Method only reads content of the collection but does not modify it.</summary>
    Read = 1
    ''' <summary>Method can change content of the collection but does not add new elements.</summary>
    ModifyExistingContent = 2
    ''' <summary>Method can add new elements to the collection.</summary>
    UpdatedContent = ModifyExistingContent Or 4
  End Enum

  ''' <summary>
  ''' Indicates that the marked method is assertion method, i.e. it halts control flow if
  ''' one of the conditions is satisfied. To set the condition, mark one of the parameters with 
  ''' <see cref="AssertionConditionAttribute"/> attribute.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Method)>
  Public NotInheritable Class AssertionMethodAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' Indicates the condition parameter of the assertion method. The method itself should be
  ''' marked by <see cref="AssertionMethodAttribute"/> attribute. The mandatory argument of
  ''' the attribute is the assertion type.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter)>
  Public NotInheritable Class AssertionConditionAttribute
      Inherits Attribute

    Public Sub New(conditionType As AssertionConditionType)
      Me.ConditionType = conditionType
    End Sub

    Private privateConditionType As AssertionConditionType
    Public Property ConditionType() As AssertionConditionType
        Get
            Return privateConditionType
        End Get
        Private Set(value As AssertionConditionType)
            privateConditionType = value
        End Set
    End Property
  End Class

  ''' <summary>
  ''' Specifies assertion type. If the assertion method argument satisfies the condition,
  ''' then the execution continues. Otherwise, execution is assumed to be halted.
  ''' </summary>
  Public Enum AssertionConditionType
    ''' <summary>Marked parameter should be evaluated to true.</summary>
    IS_TRUE = 0
    ''' <summary>Marked parameter should be evaluated to false.</summary>
    IS_FALSE = 1
    ''' <summary>Marked parameter should be evaluated to null value.</summary>
    IS_NULL = 2
    ''' <summary>Marked parameter should be evaluated to not null value.</summary>
    IS_NOT_NULL = 3
  End Enum

  ''' <summary>
  ''' Indicates that the marked method unconditionally terminates control flow execution.
  ''' For example, it could unconditionally throw exception.
  ''' </summary>
  <Obsolete("Use [ContractAnnotation('=> halt')] instead"), AttributeUsage(AttributeTargets.Method)>
  Public NotInheritable Class TerminatesProgramAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' Indicates that method is pure LINQ method, with postponed enumeration (like Enumerable.Select,
  ''' .Where). This annotation allows inference of [InstantHandle] annotation for parameters
  ''' of delegate type by analyzing LINQ method chains.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Method)>
  Public NotInheritable Class LinqTunnelAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' Indicates that IEnumerable, passed as parameter, is not enumerated.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter)>
  Public NotInheritable Class NoEnumerationAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' Indicates that parameter is regular expression pattern.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Parameter)>
  Public NotInheritable Class RegexPatternAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' XAML attribute. Indicates the type that has <c>ItemsSource</c> property and should be treated
  ''' as <c>ItemsControl</c>-derived type, to enable inner items <c>DataContext</c> type resolve.
  ''' </summary>
  <AttributeUsage(AttributeTargets.Class)>
  Public NotInheritable Class XamlItemsControlAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' XAML attibute. Indicates the property of some <c>BindingBase</c>-derived type, that
  ''' is used to bind some item of <c>ItemsControl</c>-derived type. This annotation will
  ''' enable the <c>DataContext</c> type resolve for XAML bindings for such properties.
  ''' </summary>
  ''' <remarks>
  ''' Property should have the tree ancestor of the <c>ItemsControl</c> type or
  ''' marked with the <see cref="XamlItemsControlAttribute"/> attribute.
  ''' </remarks>
  <AttributeUsage(AttributeTargets.Property)>
  Public NotInheritable Class XamlItemBindingOfItemsControlAttribute
      Inherits Attribute

  End Class

  <AttributeUsage(AttributeTargets.Class, AllowMultiple := True)>
  Public NotInheritable Class AspChildControlTypeAttribute
      Inherits Attribute

    Public Sub New(tagName As String, controlType As Type)
      Me.TagName = tagName
      Me.ControlType = controlType
    End Sub

    Private privateTagName As String
    Public Property TagName() As String
        Get
            Return privateTagName
        End Get
        Private Set(value As String)
            privateTagName = value
        End Set
    End Property
    Private privateControlType As Type
    Public Property ControlType() As Type
        Get
            Return privateControlType
        End Get
        Private Set(value As Type)
            privateControlType = value
        End Set
    End Property
  End Class

  <AttributeUsage(AttributeTargets.Property Or AttributeTargets.Method)>
  Public NotInheritable Class AspDataFieldAttribute
      Inherits Attribute

  End Class

  <AttributeUsage(AttributeTargets.Property Or AttributeTargets.Method)>
  Public NotInheritable Class AspDataFieldsAttribute
      Inherits Attribute

  End Class

  <AttributeUsage(AttributeTargets.Property)>
  Public NotInheritable Class AspMethodPropertyAttribute
      Inherits Attribute

  End Class

  <AttributeUsage(AttributeTargets.Class, AllowMultiple := True)>
  Public NotInheritable Class AspRequiredAttributeAttribute
      Inherits Attribute

    Public Sub New(<NotNull> attribute As String)
      Me.Attribute = attribute
    End Sub

    Private privateAttribute As String
    Public Property Attribute() As String
        Get
            Return privateAttribute
        End Get
        Private Set(value As String)
            privateAttribute = value
        End Set
    End Property
  End Class

  <AttributeUsage(AttributeTargets.Property)>
  Public NotInheritable Class AspTypePropertyAttribute
      Inherits Attribute

    Private privateCreateConstructorReferences As Boolean
    Public Property CreateConstructorReferences() As Boolean
        Get
            Return privateCreateConstructorReferences
        End Get
        Private Set(value As Boolean)
            privateCreateConstructorReferences = value
        End Set
    End Property

    Public Sub New(createConstructorReferences As Boolean)
      Me.CreateConstructorReferences = createConstructorReferences
    End Sub
  End Class

  <AttributeUsage(AttributeTargets.Assembly, AllowMultiple := True)>
  Public NotInheritable Class RazorImportNamespaceAttribute
      Inherits Attribute

    Public Sub New(name As String)
      Me.Name = name
    End Sub

    Private privateName As String
    Public Property Name() As String
        Get
            Return privateName
        End Get
        Private Set(value As String)
            privateName = value
        End Set
    End Property
  End Class

  <AttributeUsage(AttributeTargets.Assembly, AllowMultiple := True)>
  Public NotInheritable Class RazorInjectionAttribute
      Inherits Attribute

    Public Sub New(type As String, fieldName As String)
      Me.Type = type
      Me.FieldName = fieldName
    End Sub

    Private privateType As String
    Public Property Type() As String
        Get
            Return privateType
        End Get
        Private Set(value As String)
            privateType = value
        End Set
    End Property
    Private privateFieldName As String
    Public Property FieldName() As String
        Get
            Return privateFieldName
        End Get
        Private Set(value As String)
            privateFieldName = value
        End Set
    End Property
  End Class

  <AttributeUsage(AttributeTargets.Method)>
  Public NotInheritable Class RazorHelperCommonAttribute
      Inherits Attribute

  End Class

  <AttributeUsage(AttributeTargets.Property)>
  Public NotInheritable Class RazorLayoutAttribute
      Inherits Attribute

  End Class

  <AttributeUsage(AttributeTargets.Method)>
  Public NotInheritable Class RazorWriteLiteralMethodAttribute
      Inherits Attribute

  End Class

  <AttributeUsage(AttributeTargets.Method)>
  Public NotInheritable Class RazorWriteMethodAttribute
      Inherits Attribute

  End Class

  <AttributeUsage(AttributeTargets.Parameter)>
  Public NotInheritable Class RazorWriteMethodParameterAttribute
      Inherits Attribute

  End Class

  ''' <summary>
  ''' Prevents the Member Reordering feature from tossing members of the marked class.
  ''' </summary>
  ''' <remarks>
  ''' The attribute must be mentioned in your member reordering patterns
  ''' </remarks>
  <AttributeUsage(AttributeTargets.All)>
  Public NotInheritable Class NoReorder
      Inherits Attribute

  End Class
End Namespace
