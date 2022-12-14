// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Azure.Core;
using Azure.Identity;

namespace Microsoft.DotNet.Interactive.MicrosoftGraph;

/// <summary>
/// Helper class to create an appropriate TokenCredential
/// based on the requested authentication flow.
/// </summary>
public class CredentialProvider
{
    /// <summary>
    /// Gets the TokenCredential based on the requested authentication flow.
    /// </summary>
    /// <param name="authenticationFlow">The requested authentication flow.</param>
    /// <param name="credentialOptions">The client ID, client secret, and tenant ID from the app registration in Azure portal.</param>
    /// <param name="nationalCloud">The national cloud for authentication and Microsoft Graph service root endpoint.</param>
    /// <returns>The requested TokenCredential.</returns>
    /// <exception cref="ArgumentOutOfRangeException">The requested authentication flow was invalid.</exception>
    public static TokenCredential GetTokenCredential(
        AuthenticationFlow authenticationFlow,
        CredentialOptions credentialOptions,
        NationalCloud nationalCloud) => authenticationFlow switch
        {
            AuthenticationFlow.ClientCredential => GetClientSecretCredential(credentialOptions.TenantId, credentialOptions.ClientId, credentialOptions.ClientSecret, nationalCloud),
            AuthenticationFlow.DeviceCode => GetDeviceCodeCredential(credentialOptions.TenantId, credentialOptions.ClientId, nationalCloud),
            AuthenticationFlow.InteractiveBrowser => GetInteractiveBrowserCredential(credentialOptions.TenantId, credentialOptions.ClientId, nationalCloud),
            _ => throw new ArgumentOutOfRangeException(
                nameof(authenticationFlow),
                $"Unexpected authenticationFlow value: {authenticationFlow}"),
        };

    /// <summary>
    /// Gets the Uri for Azure Authority Host based on the nationalCloud.
    /// </summary>
    /// <param name="nationalCloud">The national cloud for authentication and Microsoft Graph service root endpoint.</param>
    /// <returns>The requested Uri for Azure Authority Host.</returns>
    /// <exception cref="ArgumentOutOfRangeException">The requested national cloud was invalid.</exception>
    public static Uri GetAuthorityHosts(
        NationalCloud nationalCloud) => nationalCloud switch
        {
            NationalCloud.Global => AzureAuthorityHosts.AzurePublicCloud,
            NationalCloud.USGov => AzureAuthorityHosts.AzureGovernment,
            NationalCloud.USGovDoD => AzureAuthorityHosts.AzureGovernment,
            _ => throw new ArgumentOutOfRangeException(
                nameof(nationalCloud),
                $"Unexpected nationalCloud value: {nationalCloud}"),
        };

    private static ClientSecretCredential GetClientSecretCredential(
        string tenantId, string clientId, string clientSecret, NationalCloud nationalCloud)
    {
        var options = new TokenCredentialOptions
        {
            AuthorityHost = GetAuthorityHosts(nationalCloud),
        };

        return new ClientSecretCredential(tenantId, clientId, clientSecret, options);
    }

    private static DeviceCodeCredential GetDeviceCodeCredential(
        string tenantId, string clientId, NationalCloud nationalCloud)
    {
        var options = new TokenCredentialOptions
        {
            AuthorityHost = GetAuthorityHosts(nationalCloud),
        };

        Func<DeviceCodeInfo, CancellationToken, Task> callback = (code, cancellation) =>
        {
            Console.WriteLine(code.Message);
            return Task.FromResult(0);
        };

        return new DeviceCodeCredential(callback, tenantId, clientId, options);
    }

    private static InteractiveBrowserCredential GetInteractiveBrowserCredential(
        string tenantId, string clientId, NationalCloud nationalCloud)
    {
        var options = new InteractiveBrowserCredentialOptions
        {
            AuthorityHost = GetAuthorityHosts(nationalCloud),
            RedirectUri = new Uri("http://localhost"),
        };

        return new InteractiveBrowserCredential(tenantId, clientId, options);
    }
}
