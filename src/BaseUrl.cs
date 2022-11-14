// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;

namespace Microsoft.DotNet.Interactive.MicrosoftGraph
{
    /// <summary>
    /// Helper class to set BaseUrl for Microsoft Graph service root endpoint.
    /// </summary>
    public class BaseUrl
    {
        /// <summary>
        /// Gets the BaseUrl for Microsoft Graph service root endpoint based on national cloud and API version.
        /// </summary>
        /// <param name="nationalCloud">The national cloud for Microsoft Graph service root endpoint.</param>
        /// <param name="apiVersion">The API version for Microsoft Graph service root endpoint.</param>
        /// <returns>The requested BaseUrl for Microsoft Graph service root endpoint.</returns>
        /// <exception cref="ArgumentOutOfRangeException">The requested national cloud and API Version pair was invalid.</exception>
        public static string GetBaseUrl(
            NationalCloud nationalCloud, ApiVersion apiVersion) => (nationalCloud, apiVersion) switch
            {
                (NationalCloud.Global, ApiVersion.V1) => new string("https://graph.microsoft.com/v1.0"),
                (NationalCloud.Global, ApiVersion.Beta) => new string("https://graph.microsoft.com/beta"),
                (NationalCloud.USGov, ApiVersion.V1) => new string("https://graph.microsoft.us/v1.0"),
                (NationalCloud.USGov, ApiVersion.Beta) => new string("https://graph.microsoft.us/beta"),
                (NationalCloud.USGovDoD, ApiVersion.V1) => new string("https://dod-graph.microsoft.us/v1.0"),
                (NationalCloud.USGovDoD, ApiVersion.Beta) => new string("https://dod-graph.microsoft.us/beta"),
                _ => throw new ArgumentOutOfRangeException(
                    nameof(nationalCloud),
                    nameof(apiVersion),
                    $"Unexpected nationalCloud and apiVersion pair values: {nationalCloud}, {apiVersion}"),
            };
    }
}
