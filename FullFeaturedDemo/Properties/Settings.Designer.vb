﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Namespace Properties


	<System.Runtime.CompilerServices.CompilerGeneratedAttribute> _
	<System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "14.0.0.0")> _
	Friend NotInheritable Partial Class Settings
		Inherits Global.System.Configuration.ApplicationSettingsBase

		Private Shared defaultInstance As Settings = DirectCast(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New Settings()), Settings)

		Public Shared ReadOnly Property [Default]() As Settings
			Get
				Return defaultInstance
			End Get
		End Property

		<System.Configuration.UserScopedSettingAttribute> _
		<System.Diagnostics.DebuggerNonUserCodeAttribute> _
		<System.Configuration.DefaultSettingValueAttribute("False")> _
		Public Property IsMaximized() As Boolean
			Get
				Return CBool(Me("IsMaximized"))
			End Get
			Set
				Me("IsMaximized") = value
			End Set
		End Property

		<System.Configuration.UserScopedSettingAttribute> _
		<System.Diagnostics.DebuggerNonUserCodeAttribute> _
		<System.Configuration.DefaultSettingValueAttribute("True")> _
		Public Property CallUpgrade() As Boolean
			Get
				Return CBool(Me("CallUpgrade"))
			End Get
			Set
				Me("CallUpgrade") = value
			End Set
		End Property

		<System.Configuration.UserScopedSettingAttribute> _
		<System.Diagnostics.DebuggerNonUserCodeAttribute> _
		<System.Configuration.DefaultSettingValueAttribute("Auto")> _
		Public Property Language() As String
			Get
				Return DirectCast(Me("Language"), String)
			End Get
			Set
				Me("Language") = value
			End Set
		End Property

		<System.Configuration.UserScopedSettingAttribute> _
		<System.Diagnostics.DebuggerNonUserCodeAttribute> _
		Public Property XmlFiles() As Global.FullFeaturedDemo.ConnectionList
			Get
				Return DirectCast(Me("XmlFiles"), Global.FullFeaturedDemo.ConnectionList)
			End Get
			Set
				Me("XmlFiles") = value
			End Set
		End Property

		<System.Configuration.UserScopedSettingAttribute> _
		<System.Diagnostics.DebuggerNonUserCodeAttribute> _
		Public Property Connections() As Global.FullFeaturedDemo.ConnectionList
			Get
				Return DirectCast(Me("Connections"), Global.FullFeaturedDemo.ConnectionList)
			End Get
			Set
				Me("Connections") = value
			End Set
		End Property
	End Class
End Namespace
