/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */


#import <Foundation/Foundation.h>
#import <office365_exchange_sdk/office365_exchange_sdk.h>

@interface Office365Snippets : NSObject

// Mail
- (void)fetchMailMessages:(void(^)(NSArray *messages, NSError *error))completion;
- (void)sendMailMessage:(MSOutlookMessage *)message
             completion:(void (^)(BOOL success, NSError *error))completion;

- (void)createDraftMailMessage:(MSOutlookMessage *)message
                    completion:(void (^)(MSOutlookMessage *addedMessage, NSError *error))completion;
- (void)createAndSendHTMLMailMessage:(NSMutableArray *)toRecipients
                          completion:(void (^)(BOOL success, NSError *error))completion;
- (void)updateMailMessage:(MSOutlookMessage *)message
               completion:(void (^)(MSOutlookMessage *updatedMessage, NSError *error))completion;
- (void)deleteMailMessage:(MSOutlookMessage *)message
               completion:(void (^)(BOOL success, NSError *error))completion;
- (void)replyToMailMessage:(MSOutlookMessage*)message
                completion:(void (^)(int success, NSError *error))completion;
- (void)createDraftReplyMessage:(MSOutlookMessage*)message
                    completion:(void (^)(MSOutlookMessage *replyMessage, NSError *error))completion;

- (void)copyMessage:(NSString*)messageId
         completion:(void (^)(BOOL success, MSODataException *error))completion;
- (void)moveMessage:(NSString*)messageId
         completion:(void (^)(BOOL success, MSODataException *error))completion;
- (void)fetchUnreadImportantMessages:(void(^)(NSArray *messages, NSError *error))completion;
- (void)fetchMessageWebLink:(void(^)(NSString *webLink, NSError *error))completion;



// Calendar
- (void)fetchCalendarEvents:(void(^)(NSArray *events, NSError *error))completion;
- (void)createCalendarEvent:(MSOutlookEvent *)event
                    completion:(void (^)(MSOutlookEvent *addedEvent, NSError *error))completion;
- (void)updateCalendarEvent:(MSOutlookEvent *)event
               completion:(void (^)(MSOutlookEvent *updatedEvent, NSError *error))completion;
- (void)deleteCalendarEvent:(MSOutlookEvent *)event
               completion:(void (^)(BOOL success, NSError *error))completion;
- (void)acceptCalendarMeetingEvent:(MSOutlookEvent *)event
                       withComment:(NSString*)comment
                        completion:(void (^)(BOOL success, NSError *error))completion;
- (void)declineCalendarMeetingEvent:(MSOutlookEvent *)event
                        withComment:(NSString*)comment
                         completion:(void (^)(BOOL success, NSError *error))completion;
- (void)tentativelyAcceptCalendarMeetingEvent:(MSOutlookEvent *)event
                                  withComment:(NSString*)comment
                                   completion:(void (^)(BOOL success, NSError *error))completion;
- (void)fetchCalendarViewFrom:(NSDate*) start
                           To:(NSDate*) end
                   completion:(void(^)(NSArray *events, NSError *error))completion;

// Contacts
- (void)fetchContacts:(void(^)(NSArray *contacts, NSError *error))completion;
- (void)createContact:(MSOutlookContact *)contact
                 completion:(void (^)(MSOutlookContact *addedContact, NSError *error))completion;
- (void)updateContact:(MSOutlookContact *)contact
                 completion:(void (^)(MSOutlookContact *updatedContact, NSError *error))completion;
- (void)deleteContact:(MSOutlookContact *)contact
                 completion:(void (^)(BOOL success, NSError *error))completion;


// Files
- (void)fetchFiles:(void (^)(NSArray *files, NSError *error))completion;



//Discovery Service
-(void) fetchDiscoveryServiceEndpoints;



// Helper Methods
- (MSOutlookMessage *)outlookMessageWithProperties:(NSArray *)recipients
                                           subject:(NSString *)subject
                                              body:(NSString *)body;

- (MSOutlookEvent *)outlookEventWithProperties:(NSArray *)attendees
                                       subject:(NSString *)subject
                                          body:(NSString *)body
                                         start: (NSDate *)start
                                           end: (NSDate *)end;

- (MSOutlookContact *)outlookContactWithProperties:(NSArray *)emailAddresses
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
