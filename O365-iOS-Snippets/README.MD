# Office 365 code snippets for iOS

**Table of contents**

* [Introduction](#introduction)
* [Set up your environment](#set-up-your-environment)
* [Use CocoaPods to import the O365 iOS SDK](#use-cocoapods-to-import-the-o365-ios-sdk)
* [Register your app with Microsoft Azure](#register-your-app-with-microsoft-azure)
* [Get the Client ID and Redirect Uri into the project](#Get-the-Client-ID-and-Redirect-Uri-into-the-project)
* [Code of Interest](#code-of-interest)
* [Questions and comments](#questions-and-comments)
* [Troubleshooting](#troubleshooting)
* [Additional resources](#additional-resources)


## Introduction


The Office 365 code snippets for iOS is a repository of code snippets that demonstrate how to work with Office 365 objects like mail, calendar, contacts, and files. This project uses the **Office 365 SDK for iOS** from [Microsoft Open Technologies, Inc (MS Open Tech)](http://msopentech.com). **This project has been tested against version 0.8.3 of the SDK.** It also uses the [Azure Active Directory Authentication Library (ADAL) for Objective-C](https://github.com/AzureAD/azure-activedirectory-library-for-objc).

The main benefits of using this sample:

* Learn how the Office 365 SDK for iOS is used to make calls that target data stored in Office 365 (including Exchange Online and SharePoint Online).

* Copy and paste code snippets that you can use in your application.

* Learn how to use the Azure Active Directory Authentication Library (ADAL) for Objective-C.

## Set up your environment

To run the Office 365 code snippets for iOS you need the following:


* [Xcode](https://developer.apple.com/) from Apple.
* An Office 365 account. You can get an Office 365 account by signing up for an [Office 365 Developer site](http://msdn.microsoft.com/en-us/library/office/fp179924.aspx). This will give you access to the APIs that you can use to create apps that target Office 365 data.
* A Microsoft Azure tenant to register your application. Azure Active Directory provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).

**Important**: You will also need to ensure your Azure subscription is bound to your Office 365 tenant. To do this see the Active Directory team's blog post, [Creating and Managing Multiple Windows Azure Active Directories](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx). The section **Adding a new directory** will explain how to do this. You can also read [Set up Azure Active Directory access for your Developer Site](http://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription) for more information.


* Installation of [CocoaPods](https://cocoapods.org/) as a dependency manager. CocoaPods will allow you to pull the Office 365 and ADAL authentication dependencies into the project.

Once you have an Office 365 account, an Azure AD account that is bound to your Office 365 Developer site, you'll need to perform the following steps:

1. Install and use CocoaPods to get the Office 365 and ADAL authentication dependencies into your project. We'll show you how to do later in the section **Use CocoaPods to import the O365 iOS SDK**.
2. Register your application with Microsoft Azure, and configure the appropriate Office 365 Exchange and SharePoint permissions.
3. Enter the Azure app registration specifics (ClientID and RedirectUri) into the Office 365 iOS Snippet app.

## Use CocoaPods to import the O365 iOS SDK
Note: If you've never used CocoaPods before as a dependency manager you'll have to install it prior to getting your O365 iOS SDK dependencies into your project. If you already have it installed you may skip this installation step and move on to **Getting the Office 365 SDK for iOS dependencies in your project**.

Enter both these lines of code from the **Terminal** app on your Mac.

    sudo gem install cocoapods
    pod setup

If the install and setup were successful, you should see the message **Setup completed in Terminal**. For more information on CocoaPods, and its usage, see [CocoaPods](https://cocoapods.org/).


**Getting the Office 365 SDK for iOS dependencies in your project.**
The O365 iOS Snippet app already contains a podfile that will get the Office 365 and ADAL components (pods) into your project. It's located in the root ("Podfile"). The syntax should look something similar to the next line.


    target ‘sample podfile’ do

    pod 'ADALiOS', '~> 1.0.0'   # 1.0.0 < ver < 1.1.0
    pod 'Office365/Outlook', '~> 0.8.3'
    pod 'Office365/Discovery', '~> 0.8.3'
    pod 'Office365/Files', '~> 0.8.3'

    end


You'll simply need to navigate to the project directory in the **Terminal** (root of the project folder) and run the following command.


    pod install

Note: You should receive confirmation that these dependencies have been added to the project and that you must open the workspace instead of the project from now on in Xcode (**O365-iOS-Snippets.xcworkspace**).  If there is a syntax error in the Podfile, you will encounter an error when you run the install command.

## Register your app with Microsoft Azure
1.	Sign in to the [Azure Management Portal](https://manage.windowsazure.com), using your Azure AD credentials.
2.	Click **Active Directory** on the left menu, then select the directory for your Office 365 developer site.
3.	On the top menu, click **Applications**.
4.	Click **Add** from the bottom menu.
5.	On the **What do you want to do page**, click **Add an application my organization is developing**.
6.	On the **Tell us about your application page**, specify **O365-iOS-Snippets** for the application name and select **NATIVE CLIENT APPLICATION** for type.
7.	Click the arrow icon on the bottom-right corner of the page.
8.	On the **Application information** page, specify a **Redirect URI**, for this example, you can specify http://localhost/snippetAppProject, and then select the checkbox in the lower-right hand corner of the page. Remember this value for the below section **Getting the ClientID and RedirectUri into the project**.
9.	Once the application has been successfully added, you will be taken to the **Quick Start** page for the application. From here, select **Configure** in the top menu.
10.	Under **permissions to other applications**, add the following two permissions: Add the **Office 365 SharePoint Online application**, and select the **Edit or delete users' files permission**. Add the **Office 365 Exchange Online application**, and select the **Read and write access to users' mail**, **Send mail as a user**, and **Have full access to users' calendars permissions**.
11.	Copy the value specified for **Client ID** on the **Configure** page. Remember this value for the below section **Getting the ClientID and RedirectUri into the project**.
12.	Click **Save** in the bottom menu.


## Get the Client ID and Redirect Uri into the project

Finally you'll need to add the Client ID and Redirect Uri you recorded from the previous section **Register your app with Microsoft Azure**.

Browse the **O365-iOS-Snippets** project directory and open up the workspace (O365-iOS-Snippets.xcworkspace). In the **AuthenticationManager.m** file you'll see that the **ClientID** and **RedirectUri** values can be added to the top of the file. Supply the necessary values here:

    // You will set your application's clientId and redirect URI. You get
    // these when you register your application in Azure AD.
    static NSString * const REDIRECT_URL_STRING = @"ENTER_REDIRECT_URI_HERE";
    static NSString * const CLIENT_ID           = @"ENTER_CLIENT_ID_HERE";
    static NSString * const AUTHORITY           = @"https://login.microsoftonline.com/common";



## Code of Interest

The **common** directory in the project contains the files needed to work with Office 365 including authentication, creating Outlook and SharePoint client objects to make Office 365 API calls, and the snippets. Another file of interest is the the MasterViewController.m. Outside of some UI logic, it defines the Office 365 snippets run in this project (Get Event, Create message, etc) and calls into the snippets library (Office365Snippets.m).

**Office 365 code snippets**

The code snippets are located in these files:

Office365Snippets.h<br>
Office365Snippets.m

**Authentication with Azure AD**

The code that uses the ADAL SDK for Objective C to authenticate with Azure AD is located in the following files:

*AuthenticationManager.h*<br>
*AuthenticationManager.m*

**Outlook Services and SharePoint client creation**

The code for creating your Outlook Services and SharePoint clients is located in the following files:

*Office365ClientFetcher.h*<br>
*Office365ClientFetcher.m*

The code in these files creates the Outlook Services,  SharePoint, and Discovery client objects. The Outlook Services and SharePoint client objects perform the calls against the Office 365 Exchange and SharePoint services. The Discovery client makes calls to get service URLs for the Outlook Services and SharePoint clients. The clients use the AuthenticationManager code to get the application an access and refresh token to act on behalf of the authenticated user.
The access and refresh token will be cached. The next time a user attempts to access the service, the access token will be issued. If the access token has expired, the client will issue the refresh token to get a new access token.


**Discovery Service**

The Office 365 Discovery service code to retrieve the Exchange and SharePoint service endpoints/URLs is located in the following file:

*Office365Snippets.m (fetchDiscoveryServiceEndpoints method)*

This sample begins the authentication process automatically when it starts. The ViewDidLoad method in MasterViewController.m calls the method **fetchDiscoveryServiceEndpoints** method. The **fetchDiscoveryServiceEndpoint** method creates a Discovery client which authenticates the user.

## Questions and comments

We'd love to get your feedback on this Office 365 iOS project . You can send your questions and suggestions to us:

* In the [Issues](https://github.com/OfficeDev/O365-iOS-Snippets/issues) section of this repository.
* On [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Make sure that your questions or comments are tagged with [Office365] and [API].

<a name="troubleshooting"></a>
## Troubleshooting

With the Xcode 7.0 update, App Transport Security is enabled for simulators and devices running iOS 9. See [App Transport Security Technote](https://developer.apple.com/library/prerelease/ios/technotes/App-Transport-Security-Technote/).

For this sample we have created a temporary exception for the following domain in the plist:

- outlook.office365.com

If this exception is not included, all calls into the Office 365 API will fail in this app when deployed to an iOS 9 simulator in Xcode.


## Additional resources

* [Office 365 APIs documentation](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [File REST operations reference](http://msdn.microsoft.com/office/office365/api/files-rest-operations)
* [Calendar REST operations reference](http://msdn.microsoft.com/office/office365/api/calendar-rest-operations)
* [Mail REST operations reference](https://msdn.microsoft.com/office/office365/api/mail-rest-operations)
* [Contacts REST operations reference](https://msdn.microsoft.com/office/office365/api/contacts-rest-operations)
* [Office Dev Center](http://dev.office.com/)
* [Microsoft Azure Active Directory Authentication Library (ADAL) for iOS and OSX](https://github.com/AzureAD/azure-activedirectory-library-for-objc)
* [Office 365 SDK for iOS](https://github.com/OfficeDev/Office-365-SDK-for-iOS)
* [UI Library for Microsoft RMS SDK v4.1 for iOS](https://github.com/AzureAD/rms-sdk-ui-for-ios)
