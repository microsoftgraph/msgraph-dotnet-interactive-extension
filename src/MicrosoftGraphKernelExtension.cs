// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

extern alias BetaLib;

using System.CommandLine;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.DotNet.Interactive.Commands;
using Microsoft.DotNet.Interactive.CSharp;
using Beta = BetaLib.Microsoft.Graph.Beta;

namespace Microsoft.DotNet.Interactive.MicrosoftGraph;

/// <summary>
/// .NET Interactive magic command extension to provide
/// authenticated Microsoft Graph clients.
/// </summary>
public class MicrosoftGraphKernelExtension : IKernelExtension
{
    /// <summary>
    /// Main entry point to extension, invoked via
    /// "#!microsoftgraph".
    /// </summary>
    /// <param name="kernel">The .NET Interactive kernel the extension is loading into.</param>
    /// <returns>A completed System.Task.</returns>
    public Task OnLoadAsync(Kernel kernel)
    {
        if (kernel is not CompositeKernel cs)
        {
            return Task.CompletedTask;
        }

        var cSharpKernel = cs.ChildKernels.OfType<CSharpKernel>().FirstOrDefault();

        var tenantIdOption = new Option<string>(
            new[] { "-t", "--tenant-id" },
            description: "Directory (tenant) ID in Azure Active Directory.");
        var clientIdOption = new Option<string>(
            new[] { "-c", "--client-id" },
            description: "Application (client) ID registered in Azure Active Directory.");
        var clientSecretOption = new Option<string>(
            new[] { "-s", "--client-secret" },
            description: "Application (client) secret registered in Azure Active Directory.");
        var scopeNameOption = new Option<string>(
            new[] { "-n", "--scope-name" },
            description: "Scope name for Microsoft Graph connection.",
            getDefaultValue: () => "graphClient");
        var authenticationFlowOption = new Option<AuthenticationFlow>(
            new[] { "-a", "--authentication-flow" },
            description: "Azure Active Directory authentication flow to use.",
            getDefaultValue: () => AuthenticationFlow.InteractiveBrowser);
        var nationalCloudOption = new Option<NationalCloud>(
            new[] { "-nc", "--national-cloud" },
            description: "National cloud for authentication and Microsoft Graph service root endpoint.",
            getDefaultValue: () => NationalCloud.Global);
        var apiVersionOption = new Option<ApiVersion>(
            new[] { "-v", "--api-version" },
            description: "Microsoft Graph API version.",
            getDefaultValue: () => ApiVersion.V1);

        var graphCommand = new Command("#!microsoftgraph", "Send Microsoft Graph requests using the specified permission flow.")
        {
            tenantIdOption,
            clientIdOption,
            clientSecretOption,
            scopeNameOption,
            authenticationFlowOption,
            nationalCloudOption,
            apiVersionOption,
        };

        graphCommand.SetHandler(
            async (string tenantId, string clientId, string clientSecret, string scopeName, AuthenticationFlow authenticationFlow, NationalCloud nationalCloud, ApiVersion apiVersion) =>
            {
                var tokenCredential = CredentialProvider.GetTokenCredential(
                    authenticationFlow, tenantId, clientId, clientSecret, nationalCloud);

                var graphServiceClient = new Beta.GraphServiceClient(tokenCredential, Scopes.GetScopes(nationalCloud));
                graphServiceClient.RequestAdapter.BaseUrl = BaseUrl.GetBaseUrl(nationalCloud, apiVersion);

                await cSharpKernel.SetValueAsync(scopeName, graphServiceClient, typeof(Beta.GraphServiceClient));
                KernelInvocationContextExtensions.Display(KernelInvocationContext.Current, $"Graph client declared with name: {scopeName}");
            },
            tenantIdOption,
            clientIdOption,
            clientSecretOption,
            scopeNameOption,
            authenticationFlowOption,
            nationalCloudOption,
            apiVersionOption);

        cSharpKernel.AddDirective(graphCommand);

        cSharpKernel.DeferCommand(new SubmitCode("using Microsoft.Graph;"));

        return Task.CompletedTask;
    }
}
