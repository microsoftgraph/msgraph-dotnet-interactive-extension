// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.DotNet.Interactive.MicrosoftGraph
{
    /// <summary>
    /// Supported authentication flows.
    /// </summary>
    public enum AuthenticationFlow
    {
        /// <summary>
        /// App-only authentication with client secret.
        /// </summary>
        ClientCredential,

        /// <summary>
        /// User authentication with device code flow.
        /// </summary>
        DeviceCode,

        /// <summary>
        /// User authentication via system default browser.
        /// </summary>
        InteractiveBrowser,
    }
}
