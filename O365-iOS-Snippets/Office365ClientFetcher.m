/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */


#import "Office365ClientFetcher.h"
#import "AuthenticationManager.h"

@implementation Office365ClientFetcher

// Gets the Outlook Services client. This will authenticate a user with the service
// and get the application an access and refresh token to act on behalf of the user.
// The access and refresh token will be cached. The next time a user attempts
// to access the service, the access token will be issued. If the access token
// has expired, the client will issue the refresh token to get a new access token.
- (void)fetchOutlookClient:(void (^)(MSOutlookClient *outlookClient))callback
{

    // Get an instance of the authentication manager controller.
    AuthenticationManager *authenticationManager = [AuthenticationManager sharedInstance];

    // This is where the client will have to provide credentials the first time.
    // The client will get back the access and refresh tokens and store them
    // in a cache.

    NSUserDefaults *userDefaults = [NSUserDefaults standardUserDefaults];
    [authenticationManager acquireAuthTokenWithResourceId: [userDefaults stringForKey:@"MailResourceID"]
                           completionHandler:^(BOOL authenticated) {


        // Gets the MSOutlookClient with the URL for the Mail service.
        if(authenticated){

            callback([[MSOutlookClient alloc] initWithUrl:[userDefaults stringForKey:@"Mail"]
                                       dependencyResolver:authenticationManager.dependencyResolver]);

        }
        else{
            //Display an alert in case of an error
            dispatch_async(dispatch_get_main_queue(), ^{
                NSLog(@"Error in the authentication");
                UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error"
                                                                message:@"Authentication failed. Check the log for errors."
                                                               delegate:self
                                                      cancelButtonTitle:@"OK"
                                                      otherButtonTitles:nil];
                [alert show];

            });
        }
    }];
}


-(void) fetchSharePointClient:(void (^)(MSSharePointClient *sharePointClient))callback{

    // Get an instance of the authentication manager controller.
    AuthenticationManager* authenticationManager = [AuthenticationManager sharedInstance];

    // Get the cached resource and URL information that was returned by discovery.
    NSUserDefaults *userDefaults = [NSUserDefaults standardUserDefaults];

    // This is where the client will have to provide credentials the first time.
    // The client will get back the access and refresh tokens and store them
    // in a cache.

    [authenticationManager acquireAuthTokenWithResourceId:[userDefaults stringForKey:@"MyFilesResourceID"]
                                        completionHandler:^(BOOL authenticated) {
        if(authenticated){


            callback([[MSSharePointClient alloc] initWithUrl:[userDefaults stringForKey:@"MyFiles"] dependencyResolver:authenticationManager.dependencyResolver]);

        }

        else{
            //Display an alert in case of an error
            dispatch_async(dispatch_get_main_queue(), ^{
                NSLog(@"Error in the authentication");
                UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error"
                                                                message:@"Authentication failed. Check the log for errors."
                                                               delegate:self
                                                      cancelButtonTitle:@"OK"
                                                      otherButtonTitles:nil];
                [alert show];

            });

        }
    }];


   }



//Gets the DiscoveryClient which is used to discover the service endpoints
- (void)fetchDiscoveryClient:(void (^)(MSDiscoveryClient *discoveryClient))callback
{

    AuthenticationManager *authenticationManager = [AuthenticationManager sharedInstance];

    [authenticationManager acquireAuthTokenWithResourceId:@"https://api.office.com/discovery/"
                                        completionHandler:^(BOOL authenticated) {
        if (authenticated) {
            callback([[MSDiscoveryClient alloc] initWithUrl:@"https://api.office.com/discovery/v1.0/me/"
                                         dependencyResolver:authenticationManager.dependencyResolver]);

        }
        else {
            dispatch_async(dispatch_get_main_queue(), ^{
                NSLog(@"Error in the authentication");
                UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error"
                                                                message:@"Authentication failed. This may be because the Internet connection is offline  or perhaps the credentials are incorrect. Check the log for errors and try again."
                                                               delegate:self
                                                      cancelButtonTitle:@"OK"
                                                      otherButtonTitles:nil];
                [alert show];

            });
        }
    }];
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
