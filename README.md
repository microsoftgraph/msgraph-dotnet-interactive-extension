# Microsoft Graph extension for .NET Interactive Notebooks

[![.NET](https://github.com/microsoftgraph/msgraph-dotnet-interactive-extension/actions/workflows/dotnet.yml/badge.svg)](https://github.com/microsoftgraph/msgraph-dotnet-interactive-extension/actions/workflows/dotnet.yml) ![License.](https://img.shields.io/badge/license-MIT-green.svg)

Sample implementation of Microsoft Graph magic command / extension for .Net Interactive.

## Test notebook

### Build and import

The below commands can be used to build as a NuGet package (C#), import, and call via magic command.

```bash
rm ~/.nuget/packages/Microsoft.DotNet.Interactive.MicrosoftGraph -Force -Recurse -ErrorAction Ignore
dotnet build ./src/Microsoft.DotNet.Interactive.MicrosoftGraph.csproj
```

```csharp
#i nuget:<REPLACE_WITH_WORKING_DIRECTORY>\src\bin\Debug\
#r "nuget:Microsoft.DotNet.Interactive.MicrosoftGraph,*-*"
```

### Test extension

Display help for "microsoftgraph" magic command

```csharp
#!microsoftgraph -h
```

Instantiate new connections to Microsoft Graph (using each authentication flow), specify unique scope name for parallel use

```csharp
#!microsoftgraph --authentication-flow InteractiveBrowser --scope-name gcInteractiveBrowser --tenant-id <tenantId> --client-id <clientId>
#!microsoftgraph --authentication-flow DeviceCode --scope-name gcDeviceCode --tenant-id <tenantId> --client-id <clientId>
#!microsoftgraph --authentication-flow ClientCredential --scope-name gcClientCredential --tenant-id <tenantId> --client-id <clientId> --client-secret <clientSecret>
```

### Interactive Browser sample snippet

```csharp
var me = await gcInteractiveBrowser.Me.Request().GetAsync();
Console.WriteLine($"Me: {me.DisplayName}, {me.UserPrincipalName}");
```

### Device Code sample snippet

```csharp
var users = await gcDeviceCode.Users.Request()
.Top(5)
.Select(u => new {u.DisplayName, u.UserPrincipalName})
.GetAsync();

users.Select(u => new {u.DisplayName, u.UserPrincipalName})
```

### Client Credential sample snippet

```csharp
var queryOptions = new List<QueryOption>()
{
new QueryOption("$count", "true")
};

var applications = await gcClientCredential.Applications
.Request( queryOptions )
.Header("ConsistencyLevel","eventual")
.Top(5)
.Select(a => new {a.AppId, a.DisplayName})
.GetAsync();

applications.Select(a => new {a.AppId, a.DisplayName})
```

## Code of conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
