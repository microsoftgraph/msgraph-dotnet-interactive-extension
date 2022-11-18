// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Text.Json;

namespace Microsoft.DotNet.Interactive.MicrosoftGraph;

/// <summary>
/// Helper class to determine app registration options
/// based on passed parameters.
/// </summary>
public class CredentialOptions
{
    private static JsonSerializerOptions jsonOptions = new JsonSerializerOptions
    {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };

    /// <summary>
    /// Gets or sets the client ID registered in the Azure portal.
    /// </summary>
    public string ClientId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the client secret registered in the Azure portal.
    /// </summary>
    public string ClientSecret { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the tenant ID from the Azure portal.
    /// </summary>
    public string TenantId { get; set; } = string.Empty;

    /// <summary>
    /// Loads the configuration file (if applicable) and combines its values
    /// with the explicit parameters. Explicit parameters override values in
    /// the configuration file.
    /// </summary>
    /// <param name="tenantId">The tenant ID parameter from the magic command.</param>
    /// <param name="clientId">The client ID parameter from the magic command.</param>
    /// <param name="clientSecret">The client secret parameter from the magic command.</param>
    /// <param name="configFile">The configuration file parameter from the magic command.</param>
    /// <returns>The determined options.</returns>
    /// <exception cref="ArgumentException">Thrown if a configuration file is specified but does not exist.</exception>
    public static CredentialOptions GetCredentialOptions(string tenantId, string clientId, string? clientSecret, string? configFile)
    {
        CredentialOptions? options = null;

        if (!string.IsNullOrEmpty(configFile))
        {
            var configFileStream = File.OpenRead(configFile);
            options = JsonSerializer.Deserialize<CredentialOptions>(configFileStream, jsonOptions);
        }

        options = options ?? new CredentialOptions();

        // Passed parameters take precedence
        options.ClientId = string.IsNullOrEmpty(clientId) ? options.ClientId : clientId;
        options.ClientSecret = string.IsNullOrEmpty(clientSecret) ? options.ClientSecret : clientSecret;
        options.TenantId = string.IsNullOrEmpty(tenantId) ? options.TenantId : tenantId;

        return options;
    }
}
