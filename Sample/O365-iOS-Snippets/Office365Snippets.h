/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */


#import <Foundation/Foundation.h>
#import <office365_exchange_sdk/office365_exchange_sdk.h>

@interface Office365Snippets : NSObject

// Mail
- (void)fetchMailMessages:(void(^)(NSArray *messages, NSError *error))completion;
- (void)sendMailMessage:(MSOutlookServicesMessage *)message
             completion:(void (^)(BOOL success, NSError *error))completion;
- (void)createDraftMailMessage:(MSOutlookServicesMessage *)message
                    completion:(void (^)(MSOutlookServicesMessage *addedMessage, NSError *error))completion;
- (void)createAndSendHTMLMailMessage:(NSMutableArray *)toRecipients
                          completion:(void (^)(BOOL success, NSError *error))completion;
- (void)updateMailMessage:(MSOutlookServicesMessage *)message
               completion:(void (^)(MSOutlookServicesMessage *updatedMessage, NSError *error))completion;
- (void)deleteMailMessage:(MSOutlookServicesMessage *)message
               completion:(void (^)(BOOL success, NSError *error))completion;

// Calendar
- (void)fetchCalendarEvents:(void(^)(NSArray *events, NSError *error))completion;
- (void)createCalendarEvent:(MSOutlookServicesEvent *)event
                    completion:(void (^)(MSOutlookServicesEvent *addedEvent, NSError *error))completion;
- (void)updateCalendarEvent:(MSOutlookServicesEvent *)event
               completion:(void (^)(MSOutlookServicesEvent *updatedEvent, NSError *error))completion;
- (void)deleteCalendarEvent:(MSOutlookServicesEvent *)event
               completion:(void (^)(BOOL success, NSError *error))completion;

// Contacts
- (void)fetchContacts:(void(^)(NSArray *contacts, NSError *error))completion;
- (void)createContact:(MSOutlookServicesContact *)contact
                 completion:(void (^)(MSOutlookServicesContact *addedContact, NSError *error))completion;
- (void)updateContact:(MSOutlookServicesContact *)contact
                 completion:(void (^)(MSOutlookServicesContact *updatedContact, NSError *error))completion;
- (void)deleteContact:(MSOutlookServicesContact *)contact
                 completion:(void (^)(BOOL success, NSError *error))completion;


// Files
- (void)fetchFiles:(void (^)(NSArray *files, NSError *error))completion;



//Discovery Service
-(void) fetchDiscoveryServiceEndpoints;



// Helper Methods
- (MSOutlookServicesMessage *)outlookMessageWithProperties:(NSArray *)recipients
                                           subject:(NSString *)subject
                                              body:(NSString *)body;

- (MSOutlookServicesEvent *)outlookEventWithProperties:(NSArray *)attendees
                                       subject:(NSString *)subject
                                          body:(NSString *)body
                                         start: (NSDate *)start
                                           end: (NSDate *)end;

- (MSOutlookServicesContact *)outlookContactWithProperties:(NSArray *)emailAddresses
                                           subject:(NSString *)givenName
                                              body:(NSString *)displayName
                                           surname:(NSString *)surname
                                             title: (NSString *)title
                                      mobilePhone1: (NSString *)mobilePhone1;






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
