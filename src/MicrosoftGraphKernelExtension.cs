// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.CommandLine;
using Microsoft.DotNet.Interactive.Commands;
using Microsoft.DotNet.Interactive.CSharp;
using Microsoft.Graph;
using Beta = Microsoft.Graph.Beta;

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
        if (cSharpKernel == null)
        {
            return Task.CompletedTask;
        }

        Option<string?> clientIdOption = new(
            new[] { "--client-id", "-c" },
            description: "Application (client) ID registered in Azure Active Directory.");

        Option<string> tenantIdOption = new(
            new[] { "--tenant-id", "-t" },
            description: "Directory (tenant) ID in Azure Active Directory.",
            getDefaultValue: () => "common");

        Option<string?> clientSecretOption = new(
            new[] { "--client-secret", "-s" },
            description: "Application (client) secret registered in Azure Active Directory.");

        Option<FileInfo> configFileOption = new(
            new[] { "--config-file", "-f" },
            description: "JSON file containing any combination of tenant ID, client ID, and client secret. Values are only used if corresponding option is not passed to the magic command.");

        Option<string> scopeNameOption = new(
            new[] { "--scope-name", "-n" },
            description: "Scope name for Microsoft Graph connection.",
            getDefaultValue: () => "graphClient");

        Option<AuthenticationFlow> authenticationFlowOption = new(
            new[] { "--authentication-flow", "-a" },
            description: "Azure Active Directory authentication flow to use.",
            getDefaultValue: () => AuthenticationFlow.InteractiveBrowser);

        Option<NationalCloud> nationalCloudOption = new(
            new[] { "--national-cloud", "-nc" },
            description: "National cloud for authentication and Microsoft Graph service root endpoint.",
            getDefaultValue: () => NationalCloud.Global);

        Option<ApiVersion> apiVersionOption = new(
            new[] { "--api-version", "-v" },
            description: "Microsoft Graph API version.",
            getDefaultValue: () => ApiVersion.V1);

        Command graphCommand = new("#!microsoftgraph", "Send Microsoft Graph requests using the specified permission flow.")
        {
            clientIdOption,
            tenantIdOption,
            clientSecretOption,
            configFileOption.ExistingOnly(),
            scopeNameOption,
            authenticationFlowOption,
            nationalCloudOption,
            apiVersionOption,
        };

        graphCommand.SetHandler(
            async (CredentialOptions credentialOptions, string scopeName, AuthenticationFlow authenticationFlow, NationalCloud nationalCloud, ApiVersion apiVersion) =>
            {
                try
                {
                    credentialOptions.ValidateOptionsForFlow(authenticationFlow);
                }
                catch (AggregateException ex)
                {
                    KernelInvocationContextExtensions.DisplayStandardError(
                        KernelInvocationContext.Current,
                        $"INVALID INPUT: {ex.Message}");
                    return;
                }

                var tokenCredential = CredentialProvider.GetTokenCredential(
                    authenticationFlow, credentialOptions, nationalCloud);

                switch (apiVersion)
                {
                    case ApiVersion.V1:
                        GraphServiceClient graphServiceClient = new(tokenCredential, Scopes.GetScopes(nationalCloud));
                        graphServiceClient.RequestAdapter.BaseUrl = BaseUrl.GetBaseUrl(nationalCloud, apiVersion);
                        await cSharpKernel.SetValueAsync(scopeName, graphServiceClient, typeof(GraphServiceClient));
                        break;
                    case ApiVersion.Beta:
                        Beta.GraphServiceClient graphServiceClientBeta = new(tokenCredential, Scopes.GetScopes(nationalCloud));
                        graphServiceClientBeta.RequestAdapter.BaseUrl = BaseUrl.GetBaseUrl(nationalCloud, apiVersion);
                        await cSharpKernel.SetValueAsync(scopeName, graphServiceClientBeta, typeof(Beta.GraphServiceClient));
                        break;
                    default:
                        break;
                }

                KernelInvocationContextExtensions.Display(KernelInvocationContext.Current, $"Graph client declared with name: {scopeName}");
            },
            new CredentialOptionsBinder(clientIdOption, tenantIdOption, clientSecretOption, configFileOption),
            scopeNameOption,
            authenticationFlowOption,
            nationalCloudOption,
            apiVersionOption);

        cSharpKernel.AddDirective(graphCommand);

        cSharpKernel.DeferCommand(new SubmitCode("using Microsoft.Graph;"));

        return Task.CompletedTask;
    }
}
