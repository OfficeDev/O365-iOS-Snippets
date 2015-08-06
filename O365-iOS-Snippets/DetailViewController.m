/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

#import "DetailViewController.h"

#import "Office365Snippets.h"

#import "AuthenticationManager.h"

@interface DetailViewController ()
@end

@implementation DetailViewController

- (void)viewDidLoad {

    [super viewDidLoad];

    [_resultsWebView loadHTMLString:[self.detailItem description] baseURL:nil];

    //Check to see if the user is logged in to the app and enable/disable the Disconnect button
    id<ADTokenCacheStoring> cache = [ADAuthenticationSettings sharedInstance].defaultTokenCacheStore;

    if (cache == nil)
    {
        self.disconnectButton.enabled = NO;
    }
    else
    {
        self.disconnectButton.enabled = YES;
    }
}

// Called to disconnect the app from accessing the user's resources. Clears
// tokens from the token store by calling clearCredentials
- (IBAction)performDisconnect:(id)sender {

    // Clear the credentials, cookies, and endpoints in NSUserDefaults.
    AuthenticationManager *authenticationManager = [AuthenticationManager sharedInstance];
    [authenticationManager clearCredentials];

    // Clear the detail view.
    [_resultsWebView loadHTMLString:@"" baseURL:nil];

    // Turn off the disconnect button. It will be enabled the next time a snippet is run.
    dispatch_async(dispatch_get_main_queue(), ^{
        self.disconnectButton.enabled = NO;
    });

}

// Called everytime a snippet is run. Provides out put to detail view.
-(void)updateUI {
    [_resultsWebView loadHTMLString:[self.detailItem description] baseURL:nil];
}

- (void)didReceiveMemoryWarning {
    [super didReceiveMemoryWarning];
    // Dispose of any resources that can be recreated.
}


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
@end
