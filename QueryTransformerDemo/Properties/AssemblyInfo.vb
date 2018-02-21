'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Reflection
Imports System.Resources
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Windows

' Управление общими сведениями о сборке осуществляется с помощью 
' набора атрибутов. Измените значения этих атрибутов, чтобы изменить сведения,
' связанные со сборкой.
<Assembly: AssemblyTitle("QueryTransformerDemo")>
<Assembly: AssemblyDescription("")>
<Assembly: AssemblyConfiguration("")>
<Assembly: AssemblyCompany("")>
<Assembly: AssemblyProduct("QueryTransformerDemo")>
<Assembly: AssemblyCopyright("Copyright © 2006-2018 Active Database Software")>
<Assembly: AssemblyTrademark("")>
<Assembly: AssemblyCulture("")>

' Параметр ComVisible со значением FALSE делает типы в сборке невидимыми 
' для COM-компонентов.  Если требуется обратиться к типу в этой сборке через 
' COM, задайте атрибуту ComVisible значение TRUE для этого типа.
<Assembly: ComVisible(False)>

'Чтобы начать сборку локализованных приложений, задайте 
'<UICulture>CultureYouAreCodingWith</UICulture> в файле .csproj
'внутри <PropertyGroup>.  Например, если используется английский США
'в своих исходных файлах установите <UICulture> в en-US.  Затем отмените преобразование в комментарий
'атрибута NeutralResourceLanguage ниже.  Обновите "en-US" в
'строка внизу для обеспечения соответствия настройки UICulture в файле проекта.

'[assembly: NeutralResourcesLanguage("en-US", UltimateResourceFallbackLocation.Satellite)]


'где расположены словари ресурсов по конкретным тематикам
'(используется, если ресурс не найден на странице 
' или в словарях ресурсов приложения)
	'где расположен словарь универсальных ресурсов
	'(используется, если ресурс не найден на странице, 
	' в приложении или в каких-либо словарях ресурсов для конкретной темы)
<Assembly: ThemeInfo(ResourceDictionaryLocation.None, ResourceDictionaryLocation.SourceAssembly)>


' Сведения о версии сборки состоят из следующих четырех значений:
'
'      Основной номер версии
'      Дополнительный номер версии 
'   Номер сборки
'      Редакция
'
' Можно задать все значения или принять номера сборки и редакции по умолчанию 
' используя "*", как показано ниже:
' [assembly: AssemblyVersion("1.0.*")]
<Assembly: AssemblyVersion("1.0.0.0")>
<Assembly: AssemblyFileVersion("1.0.0.0")>
