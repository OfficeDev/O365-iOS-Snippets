/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */



#import "AuthenticationManager.h"

// You will set your application's clientId and redirect URI. You get
// these when you register your application in Azure AD.
static NSString * const REDIRECT_URL_STRING = @"ENTER_REDIRECT_URI_HERE";
static NSString * const CLIENT_ID           = @"ENTER_CLIENT_ID_HERE";
static NSString * const AUTHORITY           = @"https://login.microsoftonline.com/common";


NSString * const Office365DidConnectNotification = @"Office365DidConnectNotification";
NSString * const Office365DidDisconnectNotification = @"Office365DidDisconnectNotification";

@interface AuthenticationManager ()

@property (strong,    nonatomic) ADAuthenticationContext *authContext;
@property (readwrite, nonatomic) ADALDependencyResolver  *dependencyResolver;

@property (readonly, nonatomic) NSURL    *redirectURL;
@property (readonly, nonatomic) NSString *authority;
@property (readonly, nonatomic) NSString *clientId;




@end

@implementation AuthenticationManager

- (instancetype)init
{
    self = [super init];

    if (self) {
        
        // These are settings that you need to set based on your
        // client registration in Azure AD.
        _redirectURL = [NSURL URLWithString:REDIRECT_URL_STRING];
        _authority = AUTHORITY;
        _clientId = CLIENT_ID;
    }

    return self;
}

// Use a single authentication controller for the application.
+ (AuthenticationManager *)sharedInstance
{
    static AuthenticationManager *sharedInstance;
    static dispatch_once_t onceToken;

    dispatch_once(&onceToken, ^{
        sharedInstance = [[AuthenticationManager alloc] init];
    });

    return sharedInstance;
}

// The first time this application is run, the [ADAuthenticationContext acquireTokenWithResource]
// manager will send a request to the AUTHORITY (see the const at the top of this file) which
// will redirect you to a login page. You will provide your credentials and the response will
// contain your refresh and access tokens. The second time this application is run, and assuming
// you didn't clear your token cache, the authentication manager will use the access or refresh
// token in the cache to authenticate client requests.
// This will result in a call to the service if you need to get an access token.

- (void)acquireAuthTokenWithResourceId:(NSString *)resourceId
                     completionHandler:(void (^)(BOOL authenticated))completionBlock
{
    ADAuthenticationError *error;
    self.authContext = [ADAuthenticationContext authenticationContextWithAuthority:self.authority
                                                                             error:&error];

    [self.authContext acquireTokenWithResource:resourceId
                                      clientId:self.clientId
                                   redirectUri:self.redirectURL
                               completionBlock:^(ADAuthenticationResult *result) {
                                   if (AD_SUCCEEDED != result.status) {

                                       completionBlock(NO);
                                   }
                                   else {
                                       NSUserDefaults *userDefaults = [NSUserDefaults standardUserDefaults];
                                       [userDefaults setObject:result.tokenCacheStoreItem.userInformation.userId
                                                        forKey:@"LogInUser"];
                                       [userDefaults synchronize];

                                       self.dependencyResolver = [[ADALDependencyResolver alloc] initWithContext:(id<ADAuthenticationContext>)self.authContext
                                                                                                   andResourceId:resourceId
                                                                                                     andClientId:self.clientId
                                                                                                  andRedirectUri:self.redirectURL];
                                       
                                       //Notification for when app is connected to O365
                                       [[NSNotificationCenter defaultCenter]postNotificationName:Office365DidConnectNotification object:nil];
                                       
                                       
                                       completionBlock(YES);
                                   }
                               }];

}

// Clear the ADAL token cache and remove this application's cookies.

-(void)clearCredentials{
    
    //Notification for when app is disconnected from O365
    [[NSNotificationCenter defaultCenter]postNotificationName:Office365DidDisconnectNotification object:nil];
    
    id<ADTokenCacheStoring> cache = [ADAuthenticationSettings sharedInstance].defaultTokenCacheStore;
    ADAuthenticationError *error;
    
    // Clear the token cache.
    if ([[cache allItemsWithError:&error] count] > 0)
        [cache removeAllWithError:&error];

    // Remove all the cookies from this application's sandbox. ADAL will try to
    // get to access tokens based on auth code in the cookie.
    NSHTTPCookieStorage *cookieStore = [NSHTTPCookieStorage sharedHTTPCookieStorage];
    for (NSHTTPCookie *cookie in cookieStore.cookies) {
        [cookieStore deleteCookie:cookie];
    }
}

@end

// *********************************************************
//
// O365-iOS-Snippets, https://github.com/OfficeDev/O365-iOS-Snippets
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************
