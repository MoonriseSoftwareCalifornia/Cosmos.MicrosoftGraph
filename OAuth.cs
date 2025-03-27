﻿// <copyright file="OAuth.cs" company="Moonrise Software, LLC">
// Copyright (c) Moonrise Software, LLC. All rights reserved.
// Licensed under the GNU Public License, Version 3.0 (https://www.gnu.org/licenses/gpl-3.0.html)
// See https://github.com/MoonriseSoftwareCalifornia/CosmosCMS
// for more information concerning the license and the contributors participating to this project.
// </copyright>

namespace Cosmos.MicrosoftGraph
{
    /// <summary>
    /// Microsoft Entra ID OAUth App Authentication.
    /// </summary>
    internal class OAuth
    {
        /// <summary>
        /// Gets or sets the client Id (Application Id) of the app.
        /// </summary>
        public string ClientId { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the Client Secret (App Secret) of the app.
        /// </summary>
        public string ClientSecret { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the Microsoft Tenant Id.
        /// </summary>
        /// <remarks>For single-tenant apps, this is the tenant id of the app registration.</remarks>
        public string TenantId { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the Callback Domain.
        /// </summary>
        /// <remarks>
        /// If you are using a proxy or firewall, such as Front Door, you may need to set this to the return domain.
        /// </remarks>
        public string CallbackDomain { get; set; } = string.Empty;

        /// <summary>
        /// Indicates if this is configured or not.
        /// </summary>
        /// <returns>True means configuration is present.</returns>
        public virtual bool IsConfigured()
        {
            return this.ClientId != string.Empty && this.ClientSecret != string.Empty;
        }
    }
}
