/** 
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

(function () {
  angular
    .module('app')
    .controller('MainController', MainController);
    
  /**
   * The MainController code.
   */
  function MainController($http, $log, adalAuthenticationService) {
		var vm = this;
		
		// Properties
		vm.isConnected;
		vm.userAlias;
		vm.emailAddress;
		vm.emailAddressSent;
		vm.requestSuccess;
		vm.requestFinished;
		vm.eventID;
		vm.eventCreated;
		
		// Methods
		vm.connect = connect;
		vm.disconnect = disconnect;
		vm.sendMail = sendMail;
		vm.createCalendarEntry = createCalendarEntry;
		vm.deleteEvent = deleteEvent;
		
		/////////////////////////////////////////
		// End of exposed properties and methods.
		
		/**
		 * This function does any initialization work the 
		 * controller needs.
		 */
    (function activate() {
			
			$log.debug('Activate Called');	

			// Check connection status and show appropriate UI.
			if (adalAuthenticationService.userInfo.isAuthenticated) {
				vm.isConnected = true;
				
				$log.debug('Authenticated.');
				
				// Get the user alias from the universal principal name (UPN).
				vm.userAlias = adalAuthenticationService.userInfo.profile.upn.split('@')[0];
				
				// Get the user email address.
				vm.emailAddress = adalAuthenticationService.userInfo.profile.upn;
			}
			else {
				$log.debug('Not Authenticated');
				vm.isConnected = false;
			}
		})();
		
		/**
		 * Expose the login method from ADAL to the view.
		 */
    function connect() {
			$log.debug('Connecting to O365....');
			adalAuthenticationService.login();			
		};
		
		/**
		 * Expose the logOut method from ADAL to the view.
		 */
    function disconnect() {
			$log.debug('Disconnecting from Office 365...');
			adalAuthenticationService.logOut();
		};
		
		/**
     * Send an email to the specified email address.
     */
    function sendMail() {
			
			//$log.debug('HTTP request to Microsoft Graph API returned successfully.', response);

			// This is the content of the email that's about to be sent.
			var emailContent = getEmailContent();
			
			// Build the HTTP request payload (the Message object).
			var email = {
				Message: {
					Subject: 'Welcome to Office 365 development with Angular and the Office 365 Connect sample',
					Body: {
						ContentType: 'HTML',
						Content: emailContent
					},
					ToRecipients: [
						{
							EmailAddress: {
								Address: vm.emailAddress
							}
						}
					]
				},
				SaveToSentItems: true
			};
			
			// Save email address so it doesn't get lost with two way data binding.
			vm.emailAddressSent = vm.emailAddress;
			
			// Build the HTTP request to send an email.
			var request = {
				method: 'POST',
				url: 'https://graph.microsoft.com/v1.0/me/microsoft.graph.sendmail',
				data: email
			};
			
			// Execute the HTTP request. 
			$http(request)
        .then(function (response) {
				$log.debug('HTTP request to Microsoft Graph API returned successfully.', response);
				response.status === 202 ? vm.requestSuccess = true : vm.requestSuccess = false;
				vm.requestFinished = true;
			}, function (error) {
				$log.error('HTTP request to Microsoft Graph API failed.');
				vm.requestSuccess = false;
				vm.requestFinished = true;
			});
		};
		
		/**
     * Gets the HTMl for the email to send.
     */
    function getEmailContent() {
			return "<html><head> <meta http-equiv=\'Content-Type\' content=\'text/html; charset=us-ascii\'> <title></title> </head><body style=\'font-family:calibri\'> <p>Congratulations " + vm.userAlias + ",</p> <p>This is a message from the Office 365 Connect sample. You are well on your way to incorporating Office 365 services in your apps. </p> <h3>What&#8217;s next?</h3> <ul><li>Check out <a href='http://dev.office.com' target='_blank'>dev.office.com</a> to start building Office 365 apps today with all the latest tools, templates, and guidance to get started quickly.</li><li>Head over to the <a href='https://msdn.microsoft.com/office/office365/howto/office-365-unified-api-reference' target='blank'>API reference on MSDN</a> to explore the rest of the APIs.</li><li>Browse other <a href='https://github.com/OfficeDev/' target='_blank'>samples on GitHub</a> to see more of the APIs in action.</li></ul> <h3>Give us feedback</h3> <ul><li>If you have any trouble running this sample, please <a href='http://github.com/OfficeDev/O365-Angular-Microsoft-Graph-Connect/issues' target='_blank'>log an issue</a>.</li><li>For general questions about the Office 365 APIs, post to <a href='http://stackoverflow.com/' target='blank'>Stack Overflow</a>. Make sure that your questions or comments are tagged with [office365].</li></ul><p>Thanks and happy coding!<br>Your Office 365 Development team </p> <div style=\'text-align:center; font-family:calibri\'> <table style=\'width:100%; font-family:calibri\'> <tbody> <tr> <td><a href=\'http://github.com/OfficeDev/O365-Angular-Microsoft-Graph-Connect'>See on GitHub</a> </td> <td><a href=\'http://officespdev.uservoice.com/'>Suggest on UserVoice</a> </td> <td><a href=\'http://twitter.com/share?text=I%20just%20started%20developing%20Angular%20apps%20using%20the%20%23Office365%20Connect%20app!%20%40OfficeDev&url=http://github.com/OfficeDev/O365-Angular-Microsoft-Graph-Connect'>Share on Twitter</a> </td> </tr> </tbody> </table> </div>  </body> </html>";
		};
		
		/**
     * Send an email to the specified email address.
     */
    function createCalendarEntry() {
			
			$log.debug('Calendar Entry Called');

			// This is the content of the calendar entry.
			var calendarContent = getCalendarContent();
			
			//Date time of the appointment (sub in the actual time)
			var StartTime = new Date();
			var EndTime = new Date();
			EndTime.setHours(EndTime.getHours() + 1);

			// Build the HTTP request payload (the Message object).
			var calendarentry = {
				subject: 'Appointment with Dr. Smith',
				body: {
					contentType: 'HTML',
					content: calendarContent
				},
				start: {
					dateTime: StartTime,
					timeZone: 'Eastern Standard Time'
				},
				end: {
					dateTime: EndTime,
					timeZone: 'Eastern Standard Time'
				},
				//attendees: [
				//	{
				//		emailAddress: {
				//			address: 'jim.smith@microsoft.com'
				//		}
				//	}					
				//],
			};
			
			$log.debug("Calendar Entry", calendarentry);

			// Save email address so it doesn't get lost with two way data binding.
			//vm.emailAddressSent = vm.emailAddress;
			
			// Build the HTTP request to create a calendar entry.
			var request = {
				method: 'POST',
				url: 'https://graph.microsoft.com/v1.0/me/events',
				data: calendarentry
			};
			
			// Execute the HTTP request. 
			$http(request)
        .then(function (response) {
				$log.debug('HTTP request to Microsoft Graph API returned successfully.', response);
				response.status === 201 ? vm.requestSuccess = true : vm.requestSuccess = false;
				vm.requestFinished = true;
				if (response.status == 201) {
					$log.debug("response.data.id = " + response.data.id);
					vm.eventID = response.data.id;
					vm.eventCreated = true;
				}
			}, function (error) {
				$log.error('HTTP request to Microsoft Graph API failed.');
				vm.requestSuccess = false;
				vm.requestFinished = true;
			});
		};
		
		/**
		 * Delete the Event
		 */
		function deleteEvent() {
			// Build the HTTP request to delete a calendar entry.
			var request = {
				method: 'DELETE',
				url: 'https://graph.microsoft.com/v1.0/me/events/' + vm.eventID,
			};
			
			// Execute the HTTP request. 
			$http(request)
        .then(function (response) {
				$log.debug('HTTP request to Microsoft Graph API returned successfully.', response);
				response.status === 204 ? vm.requestSuccess = true : vm.requestSuccess = false;
				vm.requestFinished = true;
				vm.eventCreated = false;
			}, function (error) {
				$log.error('HTTP request to Microsoft Graph API failed.');
				vm.requestSuccess = false;
				vm.requestFinished = true;
			});

		}
			
	/**
	 * Gets the HTML to send a calendar message
	 */
	function getCalendarContent() {
			return "Your appointment with Dr. Smith is confirmed";
		};
	};
})();

// *********************************************************
//
// O365-Angular-Microsoft-Graph-Connect, https://github.com/OfficeDev/O365-Angular-Microsoft-Graph-Connect
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