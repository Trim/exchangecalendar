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
 * -- Global Functions for Exchange Calendar and Exchange Contacts.
 * -- For Thunderbird.
 *
 * Author: Michel Verbraak (info@1st-setup.nl)
 * Website: http://www.1st-setup.nl/wordpress/?page_id=133
 * email: info@1st-setup.nl
 *
 *
 * ***** BEGIN LICENSE BLOCK *****/

var Cc = Components.classes;
var Ci = Components.interfaces;

var Cr = Components.results;
var components = Components;

ChromeUtils.import("resource://gre/modules/XPCOMUtils.jsm");
ChromeUtils.import("resource://gre/modules/Services.jsm");

var EXPORTED_SYMBOLS = ["exchWebService"];

if (!exchWebService) var exchWebService = {};

exchWebService.commonFunctions = {

    CreateSimpleEnumerator: function _CreateSimpleEnumerator(aArray) {
        return {
            _i: 0,
            QueryInterface: XPCOMUtils.generateQI([Ci.nsISimpleEnumerator]),
            hasMoreElements: function CSE_hasMoreElements() {
                return this._i < aArray.length;
            },
            getNext: function CSE_getNext() {
                return aArray[this._i++];
            }
        };
    },

    CreateSimpleObjectEnumerator: function _CreateSimpleObjectEnumerator(aObj) {
        return {
            _i: 0,
            _keys: Object.keys(aObj),
            QueryInterface: XPCOMUtils.generateQI([Ci.nsISimpleEnumerator]),
            hasMoreElements: function CSOE_hasMoreElements() {
                return this._i < this._keys.length;
            },
            getNext: function CSOE_getNext() {
                return aObj[this._keys[this._i++]];
            }
        };
    }
};

let mivFunctions = Cc["@1st-setup.nl/global/functions;1"].getService(Ci.mivFunctions);
exchWebService.commonFunctions.doEncodeFolderSpecialChars = mivFunctions.doEncodeFolderSpecialChars;
exchWebService.commonFunctions.encodeFolderSpecialChars = mivFunctions.encodeFolderSpecialChars;
exchWebService.commonFunctions.doDecodeFolderSpecialChars = mivFunctions.doDecodeFolderSpecialChars;
exchWebService.commonFunctions.decodeFolderSpecialChars = mivFunctions.decodeFolderSpecialChars;
exchWebService.commonFunctions.ecTZService = mivFunctions.ecTZService;
exchWebService.commonFunctions.ecDefaultTimeZone = mivFunctions.ecDefaultTimeZone;
exchWebService.commonFunctions.ecUTC = mivFunctions.ecUTC;
exchWebService.commonFunctions.splitUriGetParams = mivFunctions.splitUriGetParams;
exchWebService.commonFunctions.getBranch = mivFunctions.getBranch;
exchWebService.commonFunctions.safeGetCharPref = mivFunctions.safeGetCharPref;
exchWebService.commonFunctions.safeGetBoolPref = mivFunctions.safeGetBoolPref;
exchWebService.commonFunctions.safeGetIntPref = mivFunctions.safeGetIntPref;
exchWebService.commonFunctions.getStringBundle = mivFunctions.getStringBundle;
exchWebService.commonFunctions.getString = mivFunctions.getString;
exchWebService.commonFunctions.getUUID = mivFunctions.getUUID;
exchWebService.commonFunctions.LOG = mivFunctions.LOG;
exchWebService.commonFunctions.DEBUG = mivFunctions.DEBUG;
exchWebService.commonFunctions.writeToLogFile = mivFunctions.writeToLogFile;
exchWebService.commonFunctions.WARN = mivFunctions.WARN;
exchWebService.commonFunctions.ERROR = mivFunctions.ERROR;
exchWebService.commonFunctions.STACK = mivFunctions.STACK;
exchWebService.commonFunctions.STACKshort = mivFunctions.STACKshort;
exchWebService.commonFunctions.ASSERT = mivFunctions.ASSERT;
exchWebService.commonFunctions.copyPreferences = mivFunctions.copyPreferences;
exchWebService.commonFunctions.xmlToJxon = mivFunctions.xmlToJxon;
exchWebService.commonFunctions.splitOnCharacter = mivFunctions.splitOnCharacter;
exchWebService.commonFunctions.addCalendarById = mivFunctions.addCalendarById;
exchWebService.commonFunctions.copyCalendarSettings = mivFunctions.copyCalendarSettings;
