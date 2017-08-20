/* ***** BEGIN LICENSE BLOCK *****
 * Version: GPL 3.0
 *
 * The contents of this file are subject to the General Public License
 * 3.0 (the "License"); you may not use this file except in compliance with
 * the License. You may obtain a copy of the License at
 * http://www.gnu.org/licenses/gpl.html
 *
 * Software distributed under the License is distributed on an "AS IS" basis,
 * WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
 * for the specific language governing rights and limitations under the
 * License.
 *
 * -- Exchange 2007/2010 Calendar and Tasks Provider.
 * -- For Thunderbird with the Lightning add-on.
 *
 * This work is a combination of the Storage calendar, part of the default Lightning add-on, and 
 * the "Exchange Data Provider for Lightning" add-on currently, october 2011, maintained by Simon Schubert.
 * Primarily made because the "Exchange Data Provider for Lightning" add-on is a continuation 
 * of old code and this one is build up from the ground. It still uses some parts from the 
 * "Exchange Data Provider for Lightning" project.
 *
 * Author: Michel Verbraak (info@1st-setup.nl)
 * Website: http://www.1st-setup.nl/wordpress/?page_id=133
 * email: exchangecalendar@extensions.1st-setup.nl
 *
 *
 * This code uses parts of the Microsoft Exchange Calendar Provider code on which the
 * "Exchange Data Provider for Lightning" was based.
 * The Initial Developer of the Microsoft Exchange Calendar Provider Code is
 *   Andrea Bittau <a.bittau@cs.ucl.ac.uk>, University College London
 * Portions created by the Initial Developer are Copyright (C) 2009
 * the Initial Developer. All Rights Reserved.
 *
 * ***** BEGIN LICENSE BLOCK *****/

var Cu = Components.utils;

Cu.import("resource://gre/modules/XPCOMUtils.jsm");
Cu.import("resource://gre/modules/Services.jsm");

Cu.import("resource://exchangecalendar/ecFunctions.js");
Cu.import("resource://exchangecalendar/ecExchangeRequest.js");
Cu.import("resource://exchangecalendar/soapFunctions.js");

var EXPORTED_SYMBOLS = ["erGetFolderRequest"];

function erGetFolderRequest(aArgument, aListener) {
	this.argument = aArgument;
	this.serverUrl = aArgument.serverUrl;
	this.folderID = aArgument.folderID;
	this.folderBase = aArgument.folderBase;
	this.changeKey = aArgument.changeKey;
	this.listener = aListener;

	return this.execute();
}

erGetFolderRequest.prototype = {

	execute: function _execute() {
		let req = exchWebService.commonFunctions.xmlToJxon('<nsMessages:GetFolder xmlns:nsMessages="' + nsMessagesStr + '" xmlns:nsTypes="' + nsTypesStr + '"/>');
		req.addChildTag("FolderShape", "nsMessages", null).addChildTag("BaseShape", "nsTypes", "AllProperties");

		let parentFolder = makeParentFolderIds2("FolderIds", this.argument);
		req.addChildTagObject(parentFolder);

		let self = this;

		return this.sendRequest(self.argument, req, this.serverUrl)
			.then((exchangeResponse) => {
				let aExchangeRequest = exchangeResponse.exchangeRequest;
				let aResp = exchangeResponse.response;

				let aError = false;
				let aCode = 0;
				let aMsg = "";
				let aResult = undefined;

				let rm = aResp.XPath("/s:Envelope/s:Body/m:GetFolderResponse/m:ResponseMessages/m:GetFolderResponseMessage[@ResponseClass='Success' and m:ResponseCode='NoError']");

				if (rm.length > 0) {
					let calendarFolder = rm[0].XPath("/m:Folders/t:CalendarFolder");

					if (calendarFolder.length === 0) {
						calendarFolder = rm[0].XPath("/m:Folders/t:TasksFolder");
					}

					if (calendarFolder.length > 0) {
						var folderID = calendarFolder[0].getAttributeByTag("t:FolderId", "Id");
						var changeKey = calendarFolder[0].getAttributeByTag("t:FolderId", "ChangeKey");
						var folderClass = calendarFolder[0].getTagValue("t:FolderClass");
						self.displayName = calendarFolder[0].getTagValue("t:DisplayName");
					}
					else {
						aMsg = "Did not find any CalendarFolder parts.";
						aCode = ExchangeRequest.ER_ERROR_FINDFOLDER_FOLDERID_DETAILS;
						aError = true;
					}
				}
				else {
					aMsg = aExchangeRequest.getSoapErrorMsg(aResp);

					if (aMsg) {
						aCode = ExchangeRequest.ER_ERROR_FINDFOLDER_FOLDERID_DETAILS;
						aError = true;
					}
					else {
						aMsg = "Wrong response received.";
						aCode = ExchangeRequest.ER_ERROR_SOAP_RESPONSECODE_NOTFOUND;
						aError = true;
					}
				}

				if (aError) {
					return Promise.reject({
						exchangeRequest: aExchangeRequest,
						errorCode: aCode,
						errorMessage: aMesg
					});
				}
				else {
					return Promise.resolve({
						folderId: folderID,
						changeKey: changeKey,
						folderClass: folderClass
					});
				}
			})
			.catch((aExchangeError) => {
				return Promise.reject(aExchangeError);
			});
	},

	/* Callback encapsulation inside Promise
	 * That's only temporary needed up we use Promise for ExchangeRequest too
	 */
	sendRequest: function _sendRequest(aArgument, aXMLData, aServerURL) {
		return new Promise(function (resolve, reject) {
			let ecRequest = new ExchangeRequest(aArgument, resolve, reject);
			ecRequest.xml2json = false;
			ecRequest.xml2jxon = true;
			ecRequest.isPromise = true;
			ecRequest.sendRequest(ecRequest.makeSoapMessage(aXMLData), aServerURL);
		});
	}
};
