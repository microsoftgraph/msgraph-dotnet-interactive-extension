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
    /// <param name="tenantIdIsDefault">True if the value of tenantId is the default value.</param>
    /// <returns>The determined options.</returns>
    /// <exception cref="ArgumentException">Thrown if a configuration file is specified but does not exist.</exception>
    public static CredentialOptions GetCredentialOptions(string? tenantId, string? clientId, string? clientSecret, FileInfo? configFile, bool tenantIdIsDefault)
    {
        CredentialOptions? options = null;

        if (configFile != null)
        {
            var configFileStream = configFile.OpenRead();
            options = JsonSerializer.Deserialize<CredentialOptions>(configFileStream, jsonOptions);
        }

        options = options ?? new CredentialOptions();

        // Passed parameters take precedence
        options.ClientId = string.IsNullOrWhiteSpace(clientId) ? options.ClientId : clientId;
        options.ClientSecret = string.IsNullOrWhiteSpace(clientSecret) ? options.ClientSecret : clientSecret;

        // Tenant ID has a default value "common"
        if (string.IsNullOrEmpty(tenantId) || tenantIdIsDefault)
        {
            tenantId = tenantId ?? string.Empty;

            // If the tenant ID from the parser is the default value, only use it if
            // there isn't a value in the config file
            options.TenantId = string.IsNullOrWhiteSpace(options.TenantId) ? tenantId : options.TenantId;
        }
        else
        {
            // If it is not the default value, then it wins
            options.TenantId = tenantId;
        }

        return options;
    }

    /// <summary>
    /// Validates that the required values are present for the requested authentication flow.
    /// </summary>
    /// <param name="authenticationFlow">The requested authentication flow.</param>
    /// <exception cref="AggregateException">Thrown if there are any validation errors.</exception>
    public void ValidateOptionsForFlow(AuthenticationFlow authenticationFlow)
    {
        List<Exception> exceptions = new();

        // There must be a client ID for all flows
        if (string.IsNullOrWhiteSpace(this.ClientId))
        {
            exceptions.Add(
                new ArgumentException(
                    "A client ID must be provided in the --client-id parameter or inside the JSON file provided with the --config-file parameter."));
        }

        // There must be a client secret if using client credentials flow
        if (authenticationFlow == AuthenticationFlow.ClientCredential &&
            string.IsNullOrWhiteSpace(this.ClientSecret))
        {
            exceptions.Add(
                new ArgumentException(
                    "A client secret must be provided in the --client-secret parameter or inside the JSON file provided with the --config-file parameter."));
        }

        if (exceptions.Any())
        {
            throw new AggregateException("One or more required values are missing.", exceptions);
        }
    }
}
