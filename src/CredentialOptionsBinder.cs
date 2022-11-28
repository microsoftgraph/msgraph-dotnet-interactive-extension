// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.CommandLine;
using System.CommandLine.Binding;

namespace Microsoft.DotNet.Interactive.MicrosoftGraph;

/// <summary>
/// Binder class to combine command line parameters with a configuration file to produce a
/// CredentialOptions object.
/// </summary>
public class CredentialOptionsBinder : BinderBase<CredentialOptions>
{
    private readonly Option<string?> clientIdOption;
    private readonly Option<string> tenantIdOption;
    private readonly Option<string?> clientSecretOption;
    private readonly Option<FileInfo> configFileOption;

    /// <summary>
    /// Initializes a new instance of the <see cref="CredentialOptionsBinder"/> class.
    /// </summary>
    /// <param name="clientIdOption">The <see cref="Option"/> that provides the client ID.</param>
    /// <param name="tenantIdOption">The <see cref="Option"/> that provides the tenant ID.</param>
    /// <param name="clientSecretOption">The <see cref="Option"/> that provides the client secret.</param>
    /// <param name="configFileOption">The <see cref="Option"/> that provides the config file.</param>
    public CredentialOptionsBinder(Option<string?> clientIdOption, Option<string> tenantIdOption, Option<string?> clientSecretOption, Option<FileInfo> configFileOption)
    {
        this.clientIdOption = clientIdOption;
        this.tenantIdOption = tenantIdOption;
        this.clientSecretOption = clientSecretOption;
        this.configFileOption = configFileOption;
    }

    /// <summary>
    /// Combines values from the command line with values from the config file to create
    /// an instance of the <see cref="CredentialOptions"/> class.
    /// </summary>
    /// <param name="bindingContext">The binding context that contains the required values.</param>
    /// <returns>The <see cref="CredentialOptions"/> containing the combined values.</returns>
    protected override CredentialOptions GetBoundValue(BindingContext bindingContext)
    {
        var tenantIdIsDefaultValue = bindingContext.ParseResult.FindResultFor(this.tenantIdOption)?.IsImplicit ?? true;

        var clientId = bindingContext.ParseResult.GetValueForOption(this.clientIdOption);
        var tenantId = bindingContext.ParseResult.GetValueForOption(this.tenantIdOption);
        var clientSecret = bindingContext.ParseResult.GetValueForOption(this.clientSecretOption);
        var configFile = bindingContext.ParseResult.GetValueForOption(this.configFileOption);

        return CredentialOptions.GetCredentialOptions(tenantId, clientId, clientSecret, configFile, tenantIdIsDefaultValue);
    }
}
