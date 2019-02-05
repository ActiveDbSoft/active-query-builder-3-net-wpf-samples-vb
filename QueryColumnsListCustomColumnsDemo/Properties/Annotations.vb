'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'


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
	<AttributeUsage(AttributeTargets.Method Or AttributeTargets.Parameter Or AttributeTargets.[Property] Or AttributeTargets.[Delegate] Or AttributeTargets.Field Or AttributeTargets.[Event])> _
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
	<AttributeUsage(AttributeTargets.Method Or AttributeTargets.Parameter Or AttributeTargets.[Property] Or AttributeTargets.[Delegate] Or AttributeTargets.Field Or AttributeTargets.[Event])> _
	Public NotInheritable Class NotNullAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' Can be appplied to symbols of types derived from IEnumerable as well as to symbols of Task
	''' and Lazy classes to indicate that the value of a collection item, of the Task.Result property
	''' or of the Lazy.Value property can never be null.
	''' </summary>
	<AttributeUsage(AttributeTargets.Method Or AttributeTargets.Parameter Or AttributeTargets.[Property] Or AttributeTargets.[Delegate] Or AttributeTargets.Field)> _
	Public NotInheritable Class ItemNotNullAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' Can be appplied to symbols of types derived from IEnumerable as well as to symbols of Task
	''' and Lazy classes to indicate that the value of a collection item, of the Task.Result property
	''' or of the Lazy.Value property can be null.
	''' </summary>
	<AttributeUsage(AttributeTargets.Method Or AttributeTargets.Parameter Or AttributeTargets.[Property] Or AttributeTargets.[Delegate] Or AttributeTargets.Field)> _
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
	<AttributeUsage(AttributeTargets.Constructor Or AttributeTargets.Method Or AttributeTargets.[Property] Or AttributeTargets.[Delegate])> _
	Public NotInheritable Class StringFormatMethodAttribute
		Inherits Attribute
		''' <param name="formatParameterName">
		''' Specifies which parameter of an annotated method should be treated as format-string
		''' </param>
		Public Sub New(formatParameterName__1 As String)
			FormatParameterName = formatParameterName__1
		End Sub

		Public Property FormatParameterName() As String
			Get
				Return m_FormatParameterName
			End Get
			Private Set
				m_FormatParameterName = Value
			End Set
		End Property
		Private m_FormatParameterName As String
	End Class

	''' <summary>
	''' For a parameter that is expected to be one of the limited set of values.
	''' Specify fields of which type should be used as values for this parameter.
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.[Property] Or AttributeTargets.Field)> _
	Public NotInheritable Class ValueProviderAttribute
		Inherits Attribute
		Public Sub New(name__1 As String)
			Name = name__1
		End Sub

		<NotNull> _
		Public Property Name() As String
			Get
				Return m_Name
			End Get
			Private Set
				m_Name = Value
			End Set
		End Property
		Private m_Name As String
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
	<AttributeUsage(AttributeTargets.Parameter)> _
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
	<AttributeUsage(AttributeTargets.Method)> _
	Public NotInheritable Class NotifyPropertyChangedInvocatorAttribute
		Inherits Attribute
		Public Sub New()
		End Sub
		Public Sub New(parameterName__1 As String)
			ParameterName = parameterName__1
		End Sub

		Public Property ParameterName() As String
			Get
				Return m_ParameterName
			End Get
			Private Set
				m_ParameterName = Value
			End Set
		End Property
		Private m_ParameterName As String
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
	<AttributeUsage(AttributeTargets.Method, AllowMultiple := True)> _
	Public NotInheritable Class ContractAnnotationAttribute
		Inherits Attribute
		Public Sub New(<NotNull> contract As String)
			Me.New(contract, False)
		End Sub

		Public Sub New(<NotNull> contract__1 As String, forceFullStates__2 As Boolean)
			Contract = contract__1
			ForceFullStates = forceFullStates__2
		End Sub

		Public Property Contract() As String
			Get
				Return m_Contract
			End Get
			Private Set
				m_Contract = Value
			End Set
		End Property
		Private m_Contract As String
		Public Property ForceFullStates() As Boolean
			Get
				Return m_ForceFullStates
			End Get
			Private Set
				m_ForceFullStates = Value
			End Set
		End Property
		Private m_ForceFullStates As Boolean
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
	<AttributeUsage(AttributeTargets.All)> _
	Public NotInheritable Class LocalizationRequiredAttribute
		Inherits Attribute
		Public Sub New()
			Me.New(True)
		End Sub
		Public Sub New(required__1 As Boolean)
			Required = required__1
		End Sub

		Public Property Required() As Boolean
			Get
				Return m_Required
			End Get
			Private Set
				m_Required = Value
			End Set
		End Property
		Private m_Required As Boolean
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
	<AttributeUsage(AttributeTargets.[Interface] Or AttributeTargets.[Class] Or AttributeTargets.Struct)> _
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
	<AttributeUsage(AttributeTargets.[Class], AllowMultiple := True)> _
	<BaseTypeRequired(GetType(Attribute))> _
	Public NotInheritable Class BaseTypeRequiredAttribute
		Inherits Attribute
		Public Sub New(<NotNull> baseType__1 As Type)
			BaseType = baseType__1
		End Sub

		<NotNull> _
		Public Property BaseType() As Type
			Get
				Return m_BaseType
			End Get
			Private Set
				m_BaseType = Value
			End Set
		End Property
		Private m_BaseType As Type
	End Class

	''' <summary>
	''' Indicates that the marked symbol is used implicitly (e.g. via reflection, in external library),
	''' so this symbol will not be marked as unused (as well as by other usage inspections).
	''' </summary>
	<AttributeUsage(AttributeTargets.All)> _
	Public NotInheritable Class UsedImplicitlyAttribute
		Inherits Attribute
		Public Sub New()
			Me.New(ImplicitUseKindFlags.[Default], ImplicitUseTargetFlags.[Default])
		End Sub

		Public Sub New(useKindFlags As ImplicitUseKindFlags)
			Me.New(useKindFlags, ImplicitUseTargetFlags.[Default])
		End Sub

		Public Sub New(targetFlags As ImplicitUseTargetFlags)
			Me.New(ImplicitUseKindFlags.[Default], targetFlags)
		End Sub

		Public Sub New(useKindFlags__1 As ImplicitUseKindFlags, targetFlags__2 As ImplicitUseTargetFlags)
			UseKindFlags = useKindFlags__1
			TargetFlags = targetFlags__2
		End Sub

		Public Property UseKindFlags() As ImplicitUseKindFlags
			Get
				Return m_UseKindFlags
			End Get
			Private Set
				m_UseKindFlags = Value
			End Set
		End Property
		Private m_UseKindFlags As ImplicitUseKindFlags
		Public Property TargetFlags() As ImplicitUseTargetFlags
			Get
				Return m_TargetFlags
			End Get
			Private Set
				m_TargetFlags = Value
			End Set
		End Property
		Private m_TargetFlags As ImplicitUseTargetFlags
	End Class

	''' <summary>
	''' Should be used on attributes and causes ReSharper to not mark symbols marked with such attributes
	''' as unused (as well as by other usage inspections)
	''' </summary>
	<AttributeUsage(AttributeTargets.[Class] Or AttributeTargets.GenericParameter)> _
	Public NotInheritable Class MeansImplicitUseAttribute
		Inherits Attribute
		Public Sub New()
			Me.New(ImplicitUseKindFlags.[Default], ImplicitUseTargetFlags.[Default])
		End Sub

		Public Sub New(useKindFlags As ImplicitUseKindFlags)
			Me.New(useKindFlags, ImplicitUseTargetFlags.[Default])
		End Sub

		Public Sub New(targetFlags As ImplicitUseTargetFlags)
			Me.New(ImplicitUseKindFlags.[Default], targetFlags)
		End Sub

		Public Sub New(useKindFlags__1 As ImplicitUseKindFlags, targetFlags__2 As ImplicitUseTargetFlags)
			UseKindFlags = useKindFlags__1
			TargetFlags = targetFlags__2
		End Sub

		<UsedImplicitly> _
		Public Property UseKindFlags() As ImplicitUseKindFlags
			Get
				Return m_UseKindFlags
			End Get
			Private Set
				m_UseKindFlags = Value
			End Set
		End Property
		Private m_UseKindFlags As ImplicitUseKindFlags
		<UsedImplicitly> _
		Public Property TargetFlags() As ImplicitUseTargetFlags
			Get
				Return m_TargetFlags
			End Get
			Private Set
				m_TargetFlags = Value
			End Set
		End Property
		Private m_TargetFlags As ImplicitUseTargetFlags
	End Class

	<Flags> _
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
	<Flags> _
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
	<MeansImplicitUse(ImplicitUseTargetFlags.WithMembers)> _
	Public NotInheritable Class PublicAPIAttribute
		Inherits Attribute
		Public Sub New()
		End Sub
		Public Sub New(<NotNull> comment__1 As String)
			Comment = comment__1
		End Sub

		Public Property Comment() As String
			Get
				Return m_Comment
			End Get
			Private Set
				m_Comment = Value
			End Set
		End Property
		Private m_Comment As String
	End Class

	''' <summary>
	''' Tells code analysis engine if the parameter is completely handled when the invoked method is on stack.
	''' If the parameter is a delegate, indicates that delegate is executed while the method is executed.
	''' If the parameter is an enumerable, indicates that it is enumerated while the method is executed.
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter)> _
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
	<AttributeUsage(AttributeTargets.Method)> _
	Public NotInheritable Class PureAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' Indicates that a parameter is a path to a file or a folder within a web project.
	''' Path can be relative or absolute, starting from web root (~).
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter)> _
	Public NotInheritable Class PathReferenceAttribute
		Inherits Attribute
		Public Sub New()
		End Sub
		Public Sub New(<PathReference> basePath__1 As String)
			BasePath = basePath__1
		End Sub

		Public Property BasePath() As String
			Get
				Return m_BasePath
			End Get
			Private Set
				m_BasePath = Value
			End Set
		End Property
		Private m_BasePath As String
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
	<AttributeUsage(AttributeTargets.Method)> _
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
	<AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Method, AllowMultiple := True)> _
	Public NotInheritable Class MacroAttribute
		Inherits Attribute
		''' <summary>
		''' Allows specifying a macro that will be executed for a <see cref="SourceTemplateAttribute">source template</see>
		''' parameter when the template is expanded.
		''' </summary>
		Public Property Expression() As String
			Get
				Return m_Expression
			End Get
			Set
				m_Expression = Value
			End Set
		End Property
		Private m_Expression As String

		''' <summary>
		''' Allows specifying which occurrence of the target parameter becomes editable when the template is deployed.
		''' </summary>
		''' <remarks>
		''' If the target parameter is used several times in the template, only one occurrence becomes editable;
		''' other occurrences are changed synchronously. To specify the zero-based index of the editable occurrence,
		''' use values >= 0. To make the parameter non-editable when the template is expanded, use -1.
		''' </remarks>>
		Public Property Editable() As Integer
			Get
				Return m_Editable
			End Get
			Set
				m_Editable = Value
			End Set
		End Property
		Private m_Editable As Integer

		''' <summary>
		''' Identifies the target parameter of a <see cref="SourceTemplateAttribute">source template</see> if the
		''' <see cref="MacroAttribute"/> is applied on a template method.
		''' </summary>
		Public Property Target() As String
			Get
				Return m_Target
			End Get
			Set
				m_Target = Value
			End Set
		End Property
		Private m_Target As String
	End Class

	<AttributeUsage(AttributeTargets.Assembly, AllowMultiple := True)> _
	Public NotInheritable Class AspMvcAreaMasterLocationFormatAttribute
		Inherits Attribute
		Public Sub New(format__1 As String)
			Format = format__1
		End Sub

		Public Property Format() As String
			Get
				Return m_Format
			End Get
			Private Set
				m_Format = Value
			End Set
		End Property
		Private m_Format As String
	End Class

	<AttributeUsage(AttributeTargets.Assembly, AllowMultiple := True)> _
	Public NotInheritable Class AspMvcAreaPartialViewLocationFormatAttribute
		Inherits Attribute
		Public Sub New(format__1 As String)
			Format = format__1
		End Sub

		Public Property Format() As String
			Get
				Return m_Format
			End Get
			Private Set
				m_Format = Value
			End Set
		End Property
		Private m_Format As String
	End Class

	<AttributeUsage(AttributeTargets.Assembly, AllowMultiple := True)> _
	Public NotInheritable Class AspMvcAreaViewLocationFormatAttribute
		Inherits Attribute
		Public Sub New(format__1 As String)
			Format = format__1
		End Sub

		Public Property Format() As String
			Get
				Return m_Format
			End Get
			Private Set
				m_Format = Value
			End Set
		End Property
		Private m_Format As String
	End Class

	<AttributeUsage(AttributeTargets.Assembly, AllowMultiple := True)> _
	Public NotInheritable Class AspMvcMasterLocationFormatAttribute
		Inherits Attribute
		Public Sub New(format__1 As String)
			Format = format__1
		End Sub

		Public Property Format() As String
			Get
				Return m_Format
			End Get
			Private Set
				m_Format = Value
			End Set
		End Property
		Private m_Format As String
	End Class

	<AttributeUsage(AttributeTargets.Assembly, AllowMultiple := True)> _
	Public NotInheritable Class AspMvcPartialViewLocationFormatAttribute
		Inherits Attribute
		Public Sub New(format__1 As String)
			Format = format__1
		End Sub

		Public Property Format() As String
			Get
				Return m_Format
			End Get
			Private Set
				m_Format = Value
			End Set
		End Property
		Private m_Format As String
	End Class

	<AttributeUsage(AttributeTargets.Assembly, AllowMultiple := True)> _
	Public NotInheritable Class AspMvcViewLocationFormatAttribute
		Inherits Attribute
		Public Sub New(format__1 As String)
			Format = format__1
		End Sub

		Public Property Format() As String
			Get
				Return m_Format
			End Get
			Private Set
				m_Format = Value
			End Set
		End Property
		Private m_Format As String
	End Class

	''' <summary>
	''' ASP.NET MVC attribute. If applied to a parameter, indicates that the parameter
	''' is an MVC action. If applied to a method, the MVC action name is calculated
	''' implicitly from the context. Use this attribute for custom wrappers similar to
	''' <c>System.Web.Mvc.Html.ChildActionExtensions.RenderAction(HtmlHelper, String)</c>.
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Method)> _
	Public NotInheritable Class AspMvcActionAttribute
		Inherits Attribute
		Public Sub New()
		End Sub
		Public Sub New(anonymousProperty__1 As String)
			AnonymousProperty = anonymousProperty__1
		End Sub

		Public Property AnonymousProperty() As String
			Get
				Return m_AnonymousProperty
			End Get
			Private Set
				m_AnonymousProperty = Value
			End Set
		End Property
		Private m_AnonymousProperty As String
	End Class

	''' <summary>
	''' ASP.NET MVC attribute. Indicates that a parameter is an MVC area.
	''' Use this attribute for custom wrappers similar to
	''' <c>System.Web.Mvc.Html.ChildActionExtensions.RenderAction(HtmlHelper, String)</c>.
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter)> _
	Public NotInheritable Class AspMvcAreaAttribute
		Inherits Attribute
		Public Sub New()
		End Sub
		Public Sub New(anonymousProperty__1 As String)
			AnonymousProperty = anonymousProperty__1
		End Sub

		Public Property AnonymousProperty() As String
			Get
				Return m_AnonymousProperty
			End Get
			Private Set
				m_AnonymousProperty = Value
			End Set
		End Property
		Private m_AnonymousProperty As String
	End Class

	''' <summary>
	''' ASP.NET MVC attribute. If applied to a parameter, indicates that the parameter is
	''' an MVC controller. If applied to a method, the MVC controller name is calculated
	''' implicitly from the context. Use this attribute for custom wrappers similar to
	''' <c>System.Web.Mvc.Html.ChildActionExtensions.RenderAction(HtmlHelper, String, String)</c>.
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Method)> _
	Public NotInheritable Class AspMvcControllerAttribute
		Inherits Attribute
		Public Sub New()
		End Sub
		Public Sub New(anonymousProperty__1 As String)
			AnonymousProperty = anonymousProperty__1
		End Sub

		Public Property AnonymousProperty() As String
			Get
				Return m_AnonymousProperty
			End Get
			Private Set
				m_AnonymousProperty = Value
			End Set
		End Property
		Private m_AnonymousProperty As String
	End Class

	''' <summary>
	''' ASP.NET MVC attribute. Indicates that a parameter is an MVC Master. Use this attribute
	''' for custom wrappers similar to <c>System.Web.Mvc.Controller.View(String, String)</c>.
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter)> _
	Public NotInheritable Class AspMvcMasterAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' ASP.NET MVC attribute. Indicates that a parameter is an MVC model type. Use this attribute
	''' for custom wrappers similar to <c>System.Web.Mvc.Controller.View(String, Object)</c>.
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter)> _
	Public NotInheritable Class AspMvcModelTypeAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' ASP.NET MVC attribute. If applied to a parameter, indicates that the parameter is an MVC
	''' partial view. If applied to a method, the MVC partial view name is calculated implicitly
	''' from the context. Use this attribute for custom wrappers similar to
	''' <c>System.Web.Mvc.Html.RenderPartialExtensions.RenderPartial(HtmlHelper, String)</c>.
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Method)> _
	Public NotInheritable Class AspMvcPartialViewAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' ASP.NET MVC attribute. Allows disabling inspections for MVC views within a class or a method.
	''' </summary>
	<AttributeUsage(AttributeTargets.[Class] Or AttributeTargets.Method)> _
	Public NotInheritable Class AspMvcSupressViewErrorAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' ASP.NET MVC attribute. Indicates that a parameter is an MVC display template.
	''' Use this attribute for custom wrappers similar to 
	''' <c>System.Web.Mvc.Html.DisplayExtensions.DisplayForModel(HtmlHelper, String)</c>.
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter)> _
	Public NotInheritable Class AspMvcDisplayTemplateAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' ASP.NET MVC attribute. Indicates that a parameter is an MVC editor template.
	''' Use this attribute for custom wrappers similar to
	''' <c>System.Web.Mvc.Html.EditorExtensions.EditorForModel(HtmlHelper, String)</c>.
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter)> _
	Public NotInheritable Class AspMvcEditorTemplateAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' ASP.NET MVC attribute. Indicates that a parameter is an MVC template.
	''' Use this attribute for custom wrappers similar to
	''' <c>System.ComponentModel.DataAnnotations.UIHintAttribute(System.String)</c>.
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter)> _
	Public NotInheritable Class AspMvcTemplateAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' ASP.NET MVC attribute. If applied to a parameter, indicates that the parameter
	''' is an MVC view. If applied to a method, the MVC view name is calculated implicitly
	''' from the context. Use this attribute for custom wrappers similar to
	''' <c>System.Web.Mvc.Controller.View(Object)</c>.
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Method)> _
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
	<AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.[Property])> _
	Public NotInheritable Class AspMvcActionSelectorAttribute
		Inherits Attribute
	End Class

	<AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.[Property] Or AttributeTargets.Field)> _
	Public NotInheritable Class HtmlElementAttributesAttribute
		Inherits Attribute
		Public Sub New()
		End Sub
		Public Sub New(name__1 As String)
			Name = name__1
		End Sub

		Public Property Name() As String
			Get
				Return m_Name
			End Get
			Private Set
				m_Name = Value
			End Set
		End Property
		Private m_Name As String
	End Class

	<AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Field Or AttributeTargets.[Property])> _
	Public NotInheritable Class HtmlAttributeValueAttribute
		Inherits Attribute
		Public Sub New(<NotNull> name__1 As String)
			Name = name__1
		End Sub

		<NotNull> _
		Public Property Name() As String
			Get
				Return m_Name
			End Get
			Private Set
				m_Name = Value
			End Set
		End Property
		Private m_Name As String
	End Class

	''' <summary>
	''' Razor attribute. Indicates that a parameter or a method is a Razor section.
	''' Use this attribute for custom wrappers similar to 
	''' <c>System.Web.WebPages.WebPageBase.RenderSection(String)</c>.
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter Or AttributeTargets.Method)> _
	Public NotInheritable Class RazorSectionAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' Indicates how method, constructor invocation or property access
	''' over collection type affects content of the collection.
	''' </summary>
	<AttributeUsage(AttributeTargets.Method Or AttributeTargets.Constructor Or AttributeTargets.[Property])> _
	Public NotInheritable Class CollectionAccessAttribute
		Inherits Attribute
		Public Sub New(collectionAccessType__1 As CollectionAccessType)
			CollectionAccessType = collectionAccessType__1
		End Sub

		Public Property CollectionAccessType() As CollectionAccessType
			Get
				Return m_CollectionAccessType
			End Get
			Private Set
				m_CollectionAccessType = Value
			End Set
		End Property
		Private m_CollectionAccessType As CollectionAccessType
	End Class

	<Flags> _
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
	<AttributeUsage(AttributeTargets.Method)> _
	Public NotInheritable Class AssertionMethodAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' Indicates the condition parameter of the assertion method. The method itself should be
	''' marked by <see cref="AssertionMethodAttribute"/> attribute. The mandatory argument of
	''' the attribute is the assertion type.
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter)> _
	Public NotInheritable Class AssertionConditionAttribute
		Inherits Attribute
		Public Sub New(conditionType__1 As AssertionConditionType)
			ConditionType = conditionType__1
		End Sub

		Public Property ConditionType() As AssertionConditionType
			Get
				Return m_ConditionType
			End Get
			Private Set
				m_ConditionType = Value
			End Set
		End Property
		Private m_ConditionType As AssertionConditionType
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
	<Obsolete("Use [ContractAnnotation('=> halt')] instead")> _
	<AttributeUsage(AttributeTargets.Method)> _
	Public NotInheritable Class TerminatesProgramAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' Indicates that method is pure LINQ method, with postponed enumeration (like Enumerable.Select,
	''' .Where). This annotation allows inference of [InstantHandle] annotation for parameters
	''' of delegate type by analyzing LINQ method chains.
	''' </summary>
	<AttributeUsage(AttributeTargets.Method)> _
	Public NotInheritable Class LinqTunnelAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' Indicates that IEnumerable, passed as parameter, is not enumerated.
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter)> _
	Public NotInheritable Class NoEnumerationAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' Indicates that parameter is regular expression pattern.
	''' </summary>
	<AttributeUsage(AttributeTargets.Parameter)> _
	Public NotInheritable Class RegexPatternAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' XAML attribute. Indicates the type that has <c>ItemsSource</c> property and should be treated
	''' as <c>ItemsControl</c>-derived type, to enable inner items <c>DataContext</c> type resolve.
	''' </summary>
	<AttributeUsage(AttributeTargets.[Class])> _
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
	<AttributeUsage(AttributeTargets.[Property])> _
	Public NotInheritable Class XamlItemBindingOfItemsControlAttribute
		Inherits Attribute
	End Class

	<AttributeUsage(AttributeTargets.[Class], AllowMultiple := True)> _
	Public NotInheritable Class AspChildControlTypeAttribute
		Inherits Attribute
		Public Sub New(tagName__1 As String, controlType__2 As Type)
			TagName = tagName__1
			ControlType = controlType__2
		End Sub

		Public Property TagName() As String
			Get
				Return m_TagName
			End Get
			Private Set
				m_TagName = Value
			End Set
		End Property
		Private m_TagName As String
		Public Property ControlType() As Type
			Get
				Return m_ControlType
			End Get
			Private Set
				m_ControlType = Value
			End Set
		End Property
		Private m_ControlType As Type
	End Class

	<AttributeUsage(AttributeTargets.[Property] Or AttributeTargets.Method)> _
	Public NotInheritable Class AspDataFieldAttribute
		Inherits Attribute
	End Class

	<AttributeUsage(AttributeTargets.[Property] Or AttributeTargets.Method)> _
	Public NotInheritable Class AspDataFieldsAttribute
		Inherits Attribute
	End Class

	<AttributeUsage(AttributeTargets.[Property])> _
	Public NotInheritable Class AspMethodPropertyAttribute
		Inherits Attribute
	End Class

	<AttributeUsage(AttributeTargets.[Class], AllowMultiple := True)> _
	Public NotInheritable Class AspRequiredAttributeAttribute
		Inherits Attribute
		Public Sub New(<NotNull> attribute__1 As String)
			Attribute = attribute__1
		End Sub

		Public Property Attribute() As String
			Get
				Return m_Attribute
			End Get
			Private Set
				m_Attribute = Value
			End Set
		End Property
		Private m_Attribute As String
	End Class

	<AttributeUsage(AttributeTargets.[Property])> _
	Public NotInheritable Class AspTypePropertyAttribute
		Inherits Attribute
		Public Property CreateConstructorReferences() As Boolean
			Get
				Return m_CreateConstructorReferences
			End Get
			Private Set
				m_CreateConstructorReferences = Value
			End Set
		End Property
		Private m_CreateConstructorReferences As Boolean

		Public Sub New(createConstructorReferences__1 As Boolean)
			CreateConstructorReferences = createConstructorReferences__1
		End Sub
	End Class

	<AttributeUsage(AttributeTargets.Assembly, AllowMultiple := True)> _
	Public NotInheritable Class RazorImportNamespaceAttribute
		Inherits Attribute
		Public Sub New(name__1 As String)
			Name = name__1
		End Sub

		Public Property Name() As String
			Get
				Return m_Name
			End Get
			Private Set
				m_Name = Value
			End Set
		End Property
		Private m_Name As String
	End Class

	<AttributeUsage(AttributeTargets.Assembly, AllowMultiple := True)> _
	Public NotInheritable Class RazorInjectionAttribute
		Inherits Attribute
		Public Sub New(type__1 As String, fieldName__2 As String)
			Type = type__1
			FieldName = fieldName__2
		End Sub

		Public Property Type() As String
			Get
				Return m_Type
			End Get
			Private Set
				m_Type = Value
			End Set
		End Property
		Private m_Type As String
		Public Property FieldName() As String
			Get
				Return m_FieldName
			End Get
			Private Set
				m_FieldName = Value
			End Set
		End Property
		Private m_FieldName As String
	End Class

	<AttributeUsage(AttributeTargets.Method)> _
	Public NotInheritable Class RazorHelperCommonAttribute
		Inherits Attribute
	End Class

	<AttributeUsage(AttributeTargets.[Property])> _
	Public NotInheritable Class RazorLayoutAttribute
		Inherits Attribute
	End Class

	<AttributeUsage(AttributeTargets.Method)> _
	Public NotInheritable Class RazorWriteLiteralMethodAttribute
		Inherits Attribute
	End Class

	<AttributeUsage(AttributeTargets.Method)> _
	Public NotInheritable Class RazorWriteMethodAttribute
		Inherits Attribute
	End Class

	<AttributeUsage(AttributeTargets.Parameter)> _
	Public NotInheritable Class RazorWriteMethodParameterAttribute
		Inherits Attribute
	End Class

	''' <summary>
	''' Prevents the Member Reordering feature from tossing members of the marked class.
	''' </summary>
	''' <remarks>
	''' The attribute must be mentioned in your member reordering patterns
	''' </remarks>
	<AttributeUsage(AttributeTargets.All)> _
	Public NotInheritable Class NoReorder
		Inherits Attribute
	End Class
End Namespace
