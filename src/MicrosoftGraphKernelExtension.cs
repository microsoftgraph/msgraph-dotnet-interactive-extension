// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.CommandLine;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
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

        var tenantIdOption = new Option<string>(
            new[] { "-t", "--tenant-id" },
            description: "Directory (tenant) ID in Azure Active Directory.");
        var clientIdOption = new Option<string>(
            new[] { "-c", "--client-id" },
            description: "Application (client) ID registered in Azure Active Directory.");
        var clientSecretOption = new Option<string>(
            new[] { "-s", "--client-secret" },
            description: "Application (client) secret registered in Azure Active Directory.");
        var configFileOption = new Option<string>(
            new[] { "-f", "--config-file" },
            description: "JSON file containing any combination of tenant ID, client ID, and client secret. Values are only used if corresponding option is not passed to the magic command.",
            parseArgument: result =>
            {
                if (result.Tokens.Count == 0)
                {
                    return null;
                }

                var filePath = Path.GetFullPath(result.Tokens.Single().Value);
                if (!File.Exists(filePath))
                {
                    result.ErrorMessage = "File does not exist";
                    return null;
                }

                return filePath;
            });
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
            configFileOption,
            scopeNameOption,
            authenticationFlowOption,
            nationalCloudOption,
            apiVersionOption,
        };

        graphCommand.SetHandler(
            async (string tenantId, string clientId, string clientSecret, string configFile, string scopeName, AuthenticationFlow authenticationFlow, NationalCloud nationalCloud, ApiVersion apiVersion) =>
            {
                // Combine options to get app registration details
                var credentialOptions = CredentialOptions.GetCredentialOptions(tenantId, clientId, clientSecret, configFile);

                var tokenCredential = CredentialProvider.GetTokenCredential(
                    authenticationFlow, credentialOptions, nationalCloud);

                switch (apiVersion)
                {
                    case ApiVersion.V1:
                        var graphServiceClient = new GraphServiceClient(tokenCredential, Scopes.GetScopes(nationalCloud));
                        graphServiceClient.RequestAdapter.BaseUrl = BaseUrl.GetBaseUrl(nationalCloud, apiVersion);
                        await cSharpKernel.SetValueAsync(scopeName, graphServiceClient, typeof(GraphServiceClient));
                        break;
                    case ApiVersion.Beta:
                        var graphServiceClientBeta = new Beta.GraphServiceClient(tokenCredential, Scopes.GetScopes(nationalCloud));
                        graphServiceClientBeta.RequestAdapter.BaseUrl = BaseUrl.GetBaseUrl(nationalCloud, apiVersion);
                        await cSharpKernel.SetValueAsync(scopeName, graphServiceClientBeta, typeof(Beta.GraphServiceClient));
                        break;
                    default:
                        break;
                }

                KernelInvocationContextExtensions.Display(KernelInvocationContext.Current, $"Graph client declared with name: {scopeName}");
            },
            tenantIdOption,
            clientIdOption,
            clientSecretOption,
            configFileOption,
            scopeNameOption,
            authenticationFlowOption,
            nationalCloudOption,
            apiVersionOption);

        cSharpKernel.AddDirective(graphCommand);

        cSharpKernel.DeferCommand(new SubmitCode("using Microsoft.Graph;"));

        return Task.CompletedTask;
    }
}
