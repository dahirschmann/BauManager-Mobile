window.BAUMANAGER_CONFIG = {
  clientId: "0ba7e64b-a6e6-4839-b649-afb4087f0c79",
  tenantId: "24bf4b93-0f24-48b3-895c-b75b4c4f9b2c",

  // Multi-tenant + private Microsoft accounts, matching the app registration.
  authority: "https://login.microsoftonline.com/common",

  // Add the final published HTTPS address in Entra under:
  // Authentication -> Single-page application -> Redirect URIs.
  redirectUri: window.location.origin + window.location.pathname,

  baseFolder: "BauAufmass/Mobile",
  graphScopes: ["User.Read", "Files.ReadWrite"]
};

// Compatibility with the first prototype.
window.BAUAUFMASS_CONFIG = window.BAUMANAGER_CONFIG;
