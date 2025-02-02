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
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace MSDASC
    <ComConversionLoss>
    <StructLayout(LayoutKind.Sequential, Pack:=8)>
    Public Structure _COSERVERINFO
        Public dwReserved1 As UInteger
        <MarshalAs(UnmanagedType.LPWStr)>
        Public pwszName As String
        <ComConversionLoss>
        Public pAuthInfo As IntPtr
        Public dwReserved2 As UInteger
    End Structure

    <StructLayout(LayoutKind.Explicit, Pack:=4)>
    Public Structure __MIDL_IWinTypes_0009
        <FieldOffset(0)>
        Public hInproc As Integer
        <FieldOffset(0)>
        Public hRemote As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=4)>
    Public Structure _RemotableHandle
        Public fContext As Integer
        Public u As __MIDL_IWinTypes_0009
    End Structure

    <Guid("2206CCB1-19C1-11D1-89E0-00C04FD7A829")>
    <InterfaceTypeAttribute(CType(1, Short))>
    <ComConversionLoss>
    <ComImport>
    Public Interface IDataInitialize
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub GetDataSource(
        <MarshalAs(UnmanagedType.IUnknown), [In]> pUnkOuter As Object,
        <[In]> dwClsCtx As UInteger,
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszInitializationString As String,
        <[In]> ByRef riid As Guid,
        <MarshalAs(UnmanagedType.IUnknown), [In], Out> ByRef ppDataSource As Object)
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub GetInitializationString(
        <MarshalAs(UnmanagedType.IUnknown), [In]> pDataSource As Object,
        <[In]> fIncludePassword As SByte, <Out>
        <MarshalAs(UnmanagedType.LPWStr)> ByRef ppwszInitString As String)
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub CreateDBInstance(
        <[In]> ByRef clsidProvider As Guid,
        <MarshalAs(UnmanagedType.IUnknown), [In]> pUnkOuter As Object,
        <[In]> dwClsCtx As UInteger,
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszReserved As String,
        <[In]> ByRef riid As Guid, <Out>
        <MarshalAs(UnmanagedType.IUnknown)> ByRef ppDataSource As Object)
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub RemoteCreateDBInstanceEx(
        <[In]> ByRef clsidProvider As Guid,
        <MarshalAs(UnmanagedType.IUnknown), [In]> pUnkOuter As Object,
        <[In]> dwClsCtx As UInteger,
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszReserved As String,
        <[In]> ByRef pServerInfo As _COSERVERINFO,
        <[In]> cmq As UInteger,
        <[In]> rgpIID As IntPtr, <Out>
        <MarshalAs(UnmanagedType.IUnknown)> ByRef rgpItf As Object, <Out>
        <MarshalAs(UnmanagedType.Error)> ByRef rghr As Integer)
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub LoadStringFromStorage(
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszFileName As String, <Out>
        <MarshalAs(UnmanagedType.LPWStr)> ByRef ppwszInitializationString As String)
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub WriteStringToStorage(
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszFileName As String,
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszInitializationString As String,
        <[In]> dwCreationDisposition As UInteger)
    End Interface

    <InterfaceTypeAttribute(CType(1, Short))>
    <TypeLibTypeAttribute(CType(512, Short))>
    <Guid("2206CCB0-19C1-11D1-89E0-00C04FD7A829")>
    <ComImport>
    Public Interface IDBPromptInitialize
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub PromptDataSource(
        <MarshalAs(UnmanagedType.IUnknown), [In]> pUnkOuter As Object,
        <ComAliasName("MSDASC.wireHWND"), [In]> ByRef hWndParent As _RemotableHandle,
        <[In]> dwPromptOptions As UInteger,
        <[In]> cSourceTypeFilter As UInteger,
        <[In]> ByRef rgSourceTypeFilter As UInteger,
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszszzProviderFilter As String,
        <[In]> ByRef riid As Guid,
        <MarshalAs(UnmanagedType.IUnknown), [In], Out> ByRef ppDataSource As Object)
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub PromptFileName(
        <ComAliasName("MSDASC.wireHWND"), [In]> ByRef hWndParent As _RemotableHandle,
        <[In]> dwPromptOptions As UInteger,
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszInitialDirectory As String,
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszInitialFile As String, <Out>
        <MarshalAs(UnmanagedType.LPWStr)> ByRef ppwszSelectedFile As String)
    End Interface

    <Runtime.InteropServices.TypeLibTypeAttribute(CType(4160, Short))>
    <Guid("2206CCB2-19C1-11D1-89E0-00C04FD7A829")>
    <ComImport>
    Public Interface IDataSourceLocator
        <DispId(1610743808)>
        Property hWnd As Integer
        <DispId(1610743810)>
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Function PromptNew() As Object
        <DispId(1610743811)>
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Function PromptEdit(
        <MarshalAs(UnmanagedType.IDispatch), [In], Out> ByRef ppADOConnection As Object) As Boolean
    End Interface

    <CoClass(GetType(DataLinksClass))>
    <Guid("2206CCB2-19C1-11D1-89E0-00C04FD7A829")>
    <ComImport>
    Public Interface DataLinks
        Inherits IDataSourceLocator
    End Interface

    <TypeLibType(CType(2, Short))>
    <ComConversionLoss>
    <Guid("2206CDB2-19C1-11D1-89E0-00C04FD7A829")>
    <ComImport>
    Public Class DataLinksClass
        Implements IDataSourceLocator, DataLinks, IDBPromptInitialize, IDataInitialize

        '[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        'public extern DataLinksClass();

        <DispId(1610743808)>
        Public Overridable Property hWnd As Integer Implements IDataSourceLocator.hWnd

        <DispId(1610743810)>
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Public Overridable Function PromptNew() As Object Implements IDataSourceLocator.PromptNew
        End Function

        <DispId(1610743811)>
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Public Overridable Function PromptEdit(
        <MarshalAs(UnmanagedType.IDispatch), [In], Out> ByRef ppADOConnection As Object) As Boolean Implements IDataSourceLocator.PromptEdit
        End Function

        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Public Overridable Sub PromptDataSource(
        <MarshalAs(UnmanagedType.IUnknown), [In]> pUnkOuter As Object,
        <ComAliasName("MSDASC.wireHWND"), [In]> ByRef hWndParent As _RemotableHandle,
        <[In]> dwPromptOptions As UInteger,
        <[In]> cSourceTypeFilter As UInteger,
        <[In]> ByRef rgSourceTypeFilter As UInteger,
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszszzProviderFilter As String,
        <[In]> ByRef riid As Guid,
        <MarshalAs(UnmanagedType.IUnknown), [In], Out> ByRef ppDataSource As Object) Implements IDBPromptInitialize.PromptDataSource
        End Sub

        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Public Overridable Sub PromptFileName(
        <ComAliasName("MSDASC.wireHWND"), [In]> ByRef hWndParent As _RemotableHandle,
        <[In]> dwPromptOptions As UInteger,
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszInitialDirectory As String,
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszInitialFile As String, <Out>
        <MarshalAs(UnmanagedType.LPWStr)> ByRef ppwszSelectedFile As String) Implements IDBPromptInitialize.PromptFileName
        End Sub

        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Public Overridable Sub GetDataSource(
        <MarshalAs(UnmanagedType.IUnknown), [In]> pUnkOuter As Object,
        <[In]> dwClsCtx As UInteger,
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszInitializationString As String,
        <[In]> ByRef riid As Guid,
        <MarshalAs(UnmanagedType.IUnknown), [In], Out> ByRef ppDataSource As Object) Implements IDataInitialize.GetDataSource
        End Sub

        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Public Overridable Sub GetInitializationString(
        <MarshalAs(UnmanagedType.IUnknown), [In]> pDataSource As Object,
        <[In]> fIncludePassword As SByte, <Out>
        <MarshalAs(UnmanagedType.LPWStr)> ByRef ppwszInitString As String) Implements IDataInitialize.GetInitializationString
        End Sub

        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Public Overridable Sub CreateDBInstance(
        <[In]> ByRef clsidProvider As Guid,
        <MarshalAs(UnmanagedType.IUnknown), [In]> pUnkOuter As Object,
        <[In]> dwClsCtx As UInteger,
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszReserved As String,
        <[In]> ByRef riid As Guid, <Out>
        <MarshalAs(UnmanagedType.IUnknown)> ByRef ppDataSource As Object) Implements IDataInitialize.CreateDBInstance
        End Sub

        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Public Overridable Sub RemoteCreateDBInstanceEx(
        <[In]> ByRef clsidProvider As Guid,
        <MarshalAs(UnmanagedType.IUnknown), [In]> pUnkOuter As Object,
        <[In]> dwClsCtx As UInteger,
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszReserved As String,
        <[In]> ByRef pServerInfo As _COSERVERINFO,
        <[In]> cmq As UInteger,
        <[In]> rgpIID As IntPtr, <Out>
        <MarshalAs(UnmanagedType.IUnknown)> ByRef rgpItf As Object, <Out>
        <MarshalAs(UnmanagedType.Error)> ByRef rghr As Integer) Implements IDataInitialize.RemoteCreateDBInstanceEx
        End Sub

        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Public Overridable Sub LoadStringFromStorage(
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszFileName As String, <Out>
        <MarshalAs(UnmanagedType.LPWStr)> ByRef ppwszInitializationString As String) Implements IDataInitialize.LoadStringFromStorage
        End Sub

        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Public Overridable Sub WriteStringToStorage(
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszFileName As String,
        <MarshalAs(UnmanagedType.LPWStr), [In]> pwszInitializationString As String,
        <[In]> dwCreationDisposition As UInteger) Implements IDataInitialize.WriteStringToStorage
        End Sub
    End Class
End Namespace
