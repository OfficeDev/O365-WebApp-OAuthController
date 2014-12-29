Office 365 APIs Single Tenant Web Application
==============================================

In the sample, an existing web app with its own way of signing in users adds the ability to call Office 365 APIs. This sample uses the OAuth 2.0 authorization code grant with confidential client and the Active Directory Authentication Library (ADAL) to obtain access tokens for the web app to call the Office 365 APIs with the user's identity.

Note: For this sample, we will work with a single-tenant Azure AD application.

## How to Run this Sample
To run this sample, you need:

1. Visual Studio 2013
2. [Office Developer Tools for Visual Studio 2013](http://aka.ms/OfficeDevToolsForVS2013)
3. [Office 365 Developer Subscription](https://portal.office.com/Signup/Signup.aspx?OfferId=6881A1CB-F4EB-4db3-9F18-388898DAF510&DL=DEVELOPERPACK&ali=1)

## Step 1: Clone or download this repository
From your Git Shell or command line:

`git clone https://github.com/OfficeDev/O365-WebApp-OAuthController.git`

## Step 2: Build the Project
1. Open the project in Visual Studio 2013.
2. Simply Build the project to restore NuGet packages.
3. Ignore any build errors for now as we will configure the project in the next steps.

## Step 3: Configure the sample
Once downloaded, open the sample in Visual Studio.

### Register Azure AD application to consume Office 365 APIs
Office 365 applications use Azure Active Directory (Azure AD) to authenticate and authorize users and applications respectively. All users, application registrations, permissions are stored in Azure AD.

Using the Office 365 API Tool for Visual Studio you can configure your web application to consume Office 365 APIs.

1. In the Solution Explorer window, **right click your project -> Add -> Connected Service**.
2. A Services Manager dialog box will appear. Choose **Office 365 -> Office 365 API** and click **Register your app**.
3. On the sign-in dialog box, enter the username and password for your Office 365 tenant.
4. After you're signed in, you will see a list of all the services.
5. Initially, no permissions will be selected, as the app is not registered to consume any services yet.
6. Select **Users and Groups** and then click **Permissions**
7. In the **Users and Groups Permissions** dialog, select **Enable sign-on and read users profiles'** and click **Apply**
8. Select **My Files** and then click **Permissions**
9. In the **My Files Permissions** dialog, select **Read users' files** and click **Apply**
10. Click **Ok**

After clicking OK in the Services Manager dialog box, Office 365 client libraries (in the form of NuGet packages) for connecting to Office 365 APIs will be added to your project.

In this process, Office 365 API tool registered an Azure AD Application in the Office 365 tenant that you signed in the wizard and added the Azure AD application details to web.config.

### Step 4: Update web.config with your Tenant ID and Tenant Name
One of the things that is required for your single tenant web application to successfully authenticate against Azure Active Directory is to specify the login authority along with the tenant Id. The authority looks like this:
https://login.windows.net/{tenantId}

In order to complete this procedure, you're going to need to log into the Microsoft Azure management portal. To do this you must have an Azure subscription. A free trial is available if you do not currently have one.
You can sign up here: http://azure.microsoft.com/en-us/pricing/free-trial/. You must also ensure you have already completed the Register app to consume Office 365 APIs procedure.

Note: You will also need to ensure your Azure subscription is bound to your Office 365 tenant. To do this see the Active Directory team's blog post, [Creating and Managing Multiple Windows Azure Active Directories](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx). The section **Adding a new directory** will explain how to do this. You can also read [Set up Azure Active Directory access for your Developer Site](http://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription) for more information.

To update the web.config with your Office 365 tenant ID and tenant name:

1. Sign into the Azure management portal at https://manage.windowsazure.com/.
2. Select the Active Directory tab in the left pane and choose your target Office 365 domain underneath the back button. As a reminder you must have the Azure subscription configured to use your specific Office 365 tenant.

![](http://i.imgur.com/SU8Ri5f.png)

3. Choose the Applications tab for your domain and select the registration entry for your app. It should appear as something like O365-APIs-Start-ASPNET-MVC.OfficeO365App.

![](http://i.imgur.com/5dtWcua.png)

4. Upon clicking that entry, expand the Enable Users To Sign On section, copy and paste the Federation Metadata Document URL value to notepad or another application. You'll notice that there's an identifier present in that URL (in the form of a guid), and this is the tenant ID that is needed for the project.

![](http://i.imgur.com/TzXIlut.png)

5. Copy just the identifier value and return to the sample solution.
6. Add your tenant ID to the ida:TenantID key in the web.config. It should look similar to this:
`<add key="ida:TenantId" value="d10f81ac-2de0-4eaf-af91-393d1bdaf17d"/>`
7. Add your tenant name to the ida:Tenant key in the web.config.

Note: If you are deploying to a production tenant, you will need to ask your tenant admin for the tenant identifier.

### Step 5: Build and Debug your web application
Now you are ready for a test run. Hit F5 to test the app.

### Quick Look at the OAuth Controller Code
The **OAuthController.cs** will be invoked as a call-back from an authentication service (e.g., https://login.windows.net/). On completion, the method will cache the refresh token and access tokens, and redirect to the URL specified in the state parameter.

### Write Code to call Office 365 APIs
You can now write code to call an Office 365 API in your web application. 

**FilesController.cs** describes how to interact with the Office 365 API discovery service, get the OneDrive endpoint URI and resource Id and then query users' files.

If the user has not consented the app, then the code constructs an authorization Url which will redirect user to login to Office 365. 
