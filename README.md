# VB.NET Demo Projects for [Active Query Builder for .NET](https://www.activequerybuilder.com/product_net.html) - WPF Edition

##### Also, check [the C# Demo projects repository](https://github.com/ActiveDbSoft/active-query-builder-3-net-wpf-samples-csharp) for the same demos in C#.
#
## What is Active Query Builder for .NET?
Active Query Builder is a component suite for different .NET environments: [WinForms](https://www.activequerybuilder.com/product_winforms.html), [WPF](https://www.activequerybuilder.com/product_wpf.html), [ASP.NET](https://www.activequerybuilder.com/product_asp.html). 
It includes:
- Visual SQL Query Builder,
- SQL Parser and Analyzer,
- API to build and modify SQL queries,
- SQL Text Editor with code completion and syntax highlighting.

##### Details:
- [Active Query Builder website](http://www.activequerybuilder.com/),
- [Active Query Builder for .NET details page](http://www.activequerybuilder.com/product_net.html).

## How do I get Active Query Builder?
- [Download the trial version](https://www.activequerybuilder.com/trequest.html?request=net) from the product web site
- Get it by installing the [Active Query Builder for .NET - WPF NuGet package](https://www.nuget.org/packages/ActiveQueryBuilder.View.WPF/).

## What's in this repository?
The demo projects in this repository illustrate various aspects of the component's functionality from basic usage scenarios to advanced features. They are also included the trial and full versions of Active Query Builder.

##### Prerequisites:
- Visual Studio 2012 or higher,
- .NET Framework 4.0 or higher.

## How to get these demo projects up and running?

1. Clone this repository to your PC.
2. Open the "**WPFDemo_VisualBasic.sln**" solution.
3. Run the project.

The necessary packages will be installed automatically. In case of any problems with the packages, open the "Tools" - "NuGet Package Manager" - "Package Manager Console" menu item and install them by running the following command: 

    Install-Package ActiveQueryBuilder.View.WPF

Some demo projects require third-party DB connection libraries to run. For those libraries which have up-to-date Nuget packages available, the necessary assemblies will be updated automatically. Assemblies for the rest of the libraries can be found in the "third-party" folder. Note that for most of those libraries, installation of the full database client is necessary. The assembly will allow only to compile the project, but not get connected to the database server.

## Have a question or want to leave feedback?

Welcome to the [Active Query Builder Help Center](https://support.activequerybuilder.com/hc/)!
There you will find:
- End-user's Guide,
- Getting Started guides,
- Knowledge Base,
- Community Forum.

## License
The source code of the demo projects in this repository is covered by the [MIT license](https://en.wikipedia.org/wiki/MIT_License).