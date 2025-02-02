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


Namespace Connection
    Public Delegate Sub SyntaxProviderDetected(syntaxType As Type)

    Friend Interface IConnectionFrame
        Event OnSyntaxProviderDetected As SyntaxProviderDetected
        Sub SetServerType(serverType As String)

        Property ConnectionString() As String
        Function TestConnection() As Boolean
    End Interface
End Namespace
