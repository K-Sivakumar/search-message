Search messages from Mail and Teams chat using MS Graph
-------------------------------------------------------
### Intro
- This is a sample application to search mail and team chat based on the given text, the search result will contain the Channel (Mail/Team chat), Received date, Received from, Subject (in case of mail) and Summary information 
### Prerequisites
- [.NET Core SDK](https://dotnet.microsoft.com/download)
### Register an app in Azure AD

1. Open a browser and navigate to the [Azure Active Directory admin center](https://aad.portal.azure.com). Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.

1. Select **Azure Active Directory** in the left-hand navigation, then select **App registrations** under **Manage**.

1. Select **New registration**. On the **Register an application** page, set the values as follows.

    - Set **Name** to `search-msg`.
    - Set **Supported account types** to `Accounts in this organizational directory only (MSFT only - Single tenant)`.
    - Under **Redirect URI**, set the first drop-down to `Web` and set the value to `https://localhost:5001/signin-oidc`.

1. Select **Register**. On the **search-msg** page, copy the values of the **Application (client) ID** and **Directory (tenant) ID** and save it, you will need it in the next step.
1. Select **Certificates & secrets** in the left-hand navigation,then select **New client secret** under **Client secrets**

    - Set **Description** to `search-msg`
    - Set **Expires** to `Recommended: 180 days (6 months)`
1. Select **Add**. On the **Client secrets (1)** page, copy the value of the **Value** field and save it, you will never view this again and will need it in the next step.
### Configure the sample
1. Edit **./appsettings.json** and replace the following code.

    ```json
    {
      "AzureAd": {
        "TenantId": "YOUR_TENANT_ID_HERE",
        "ClientId": "YOUR_CLIENT_ID_HERE",
        "ClientSecret": "YOUR_CLIENT_SECRET_HERE"
      }
    }
    ```

1. Replace `YOUR_TENANT_ID_HERE` with the **Directory (tenant) ID** value from your app registration.
1. Replace `YOUR_CLIENT_ID_HERE` with the **Application (client) ID** value from your app registration.
1. Replace `YOUR_CLIENT_SECRET_HERE` with the **Value** from your client secrets
### Running the sample

Run the following command in the **search-message** directory.

```dotnetcli
dotnet run
```

Open your browser to `https://localhost:5001`.

#### Note

As a sample, and for simplicity, this sample does not follow best practices for an application in Production. If you intent to ship this code, we recommend doing the following:

-	Never check in app secrets to your repository
-	Never use In-memory token cache

-	**NuGet packages** used in the sample. These are handled using the package manager, as described in the setup instructions. These should update automatically at build time; if not, make sure your NuGet package manager is up-to-date. You can learn more about the packages we used at the links below.

	-	[Microsoft.Identity.Web.MicrosoftGraphBeta](https://github.com/AzureAD/microsoft-identity-web) provides MS Graph API calls.
	-	[Microsoft.Identity.Web](https://github.com/AzureAD/microsoft-identity-web) provides OAuth with OpenID authentication and authorization.
	-	[Microsoft.Identity.Web.UI](https://github.com/AzureAD/microsoft-identity-web) provides authentication and authorization UI.
