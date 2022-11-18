// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.DotNet.Interactive.MicrosoftGraph;

/// <summary>
/// Helper class to set scopes for authentication based on national cloud.
/// </summary>
public class Scopes
{
    /// <summary>
    /// Gets the Uri for Azure Authority Host based on the nationalCloud.
    /// </summary>
    /// <param name="nationalCloud">The national cloud for authentication and Microsoft Graph service root endpoint.</param>
    /// <returns>The requested Uri for Azure Authority Host.</returns>
    /// <exception cref="ArgumentOutOfRangeException">The requested national cloud was invalid.</exception>
    public static string[] GetScopes(
        NationalCloud nationalCloud) => nationalCloud switch
        {
            NationalCloud.Global => new string[] { "https://graph.microsoft.com/.default" },
            NationalCloud.USGov => new string[] { "https://graph.microsoft.us/.default" },
            NationalCloud.USGovDoD => new string[] { "https://dod-graph.microsoft.us/.default" },
            _ => throw new ArgumentOutOfRangeException(
                nameof(nationalCloud),
                $"Unexpected nationalCloud value: {nationalCloud}"),
        };
}
