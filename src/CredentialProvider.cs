// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Threading;
using System.Threading.Tasks;
using Azure.Core;
using Azure.Identity;

namespace Microsoft.DotNet.Interactive.MicrosoftGraph
{
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
        /// <param name="tenantId">Tenant ID for single-tenant apps, or "common" for multi-tenant.</param>
        /// <param name="clientId">The client ID from the app registration in Azure portal.</param>
        /// <param name="clientSecret">The client secret (only applicable to ClientCredential flow).</param>
        /// <param name="nationalCloud">The national cloud for authentication and Microsoft Graph service root endpoint.</param>
        /// <returns>The requested TokenCredential.</returns>
        /// <exception cref="ArgumentOutOfRangeException">The requested authentication flow was invalid.</exception>
        public static TokenCredential GetTokenCredential(
            AuthenticationFlow authenticationFlow,
            string tenantId,
            string clientId,
            string clientSecret,
            NationalCloud nationalCloud) => authenticationFlow switch
            {
                AuthenticationFlow.ClientCredential => GetClientSecretCredential(tenantId, clientId, clientSecret, nationalCloud),
                AuthenticationFlow.DeviceCode => GetDeviceCodeCredential(tenantId, clientId, nationalCloud),
                AuthenticationFlow.InteractiveBrowser => GetInteractiveBrowserCredential(tenantId, clientId, nationalCloud),
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
}
