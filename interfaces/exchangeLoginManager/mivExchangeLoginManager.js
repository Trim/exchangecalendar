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
 * Author: Adrien Dorsaz
 * Website: https://adorsaz.ch
 * Contact: adrien@adorsaz.ch
 *
 * ***** BEGIN LICENSE BLOCK *****/

/*
 * This interface is a service providing management for exchange passwords.
 *
 * Its goals are to:
 *  - ask passwords to user when needed
 *  - keep a list of given passwords for the current session
 *    (exchange communications need permanent connection to server)
 *  - ask Mozilla Login Manager to save passwords on disk if asked
 *
 * It relies on Mozilla Login Manager interfaces to prompt passwords.
 */

var Cc = Components.classes;
var Ci = Components.interfaces;
var Cu = Components.utils;

Cu.import("resource://gre/modules/XPCOMUtils.jsm");

// Constructor
function mivExchangeLoginManager() {
    this.loginCache = [];
}

// Definition
mivExchangeLoginManager.prototype = {
    // XPCOM Properties
    classDescription: "Exchange Add-on Login Manager Service",
    classID: Components.ID("{7451198f-a392-41d3-b255-d1fdca7533b2}"),
    contractID: "@1st-setup.nl/exchange/loginmanager;1",

    flags: Ci.nsIClassInfo.SINGLETON || Ci.nsIClassInfo.THREADSAFE,
    implementationLanguage: Ci.nsIProgrammingLanguage.JAVASCRIPT,

    QueryInterface: XPCOMUtils.generateQI([Ci.mivExchangeLoginManager]),

    // Exchange Login Manager methods

    // Method to query passwords
    AUTF8String getPassword(in AUTF8String login, in AUTF8String serverURL, in AUTF8String httpRealm);
    {
        var password = null;

        this.logInfo("  --- mivExchangeLoginManager.getPassword()");

        // We'll need LoginInfo to use the Mozilla Login Manager
        var nsLoginInfo = Components.Constructor(
                            "@mozilla.org/login-manager/loginInfo;1",
                            Ci.nsILoginInfo,
                            "init");

        // Exchange calendar add-on authenticate user by HTTP, not HTML Form
        var loginInfo = new nsLoginInfo(
                                serverURL, //hostname
                                "", // action URL in HTML form (blank to be ignored)
                                httpRealm, // HTTP WWW-Authenticate Basic Realm
                                login, // User name
                                "", // Password (unknown currently)
                                "", // User name input field attribute HTML form
                                "" // Password input field attribute HTML form
                        );

        // Check if password available in session cache
        for (var i=0; i< this.loginCache.length; i++) {
            var cachedLogin = this.loginCache[i];

            if(loginInfo.matches(cachedLogin, false) and !cachedLogin.password){
                this.logInfo("  - password found in cache");

                password = cachedLogin.password;
                break;
            }
        }

        // Password not in cache, ask it to Mozilla's Login Manager
        if(!password){
            var mozLoginManager = Cc["@mozilla.org/login-manager;1"].getService(
                                    Ci.nsILoginManager);

            // Find all logins corresponding to this request
            var mozLogins = mozLoginManager.findLogins({}, serverURL, "", httpRealm);

            for (var i=0 ; i < mozLogins.length(); i++){
                if (mozLogins[i].username == loginInfo.username) {
                    this.logInfo("  - password found in Mozilla Login Manager");

                    password = mozLogins[i].password;
                    this.cachePassword(loginInfo, password);
                    break;
                }
            }
        }

        // Password not found, ask it to user
        if(!password){
            var myAuthPrompt = Cc["@mozilla.org/login-manager/prompter;1"].createInstance(
                                    Ci.nsIAuthPrompt);

            // Consctruct passwordRealm from user and url:
            var uriStart = serverURL.indexOf("://") + 3;
            var passwordRealm = this.serverURL.substr(0, uriStart) + urlencode(login)
            						+ "@" + serverURL.substr(uriStart);
            var promptTitle = "Exchange calendar password prompt";
            var promptText = "Please give your exchange password for login "
                             + login + " on server " + serverURL
                             + " (" + passwordRealm + ")";
            var passwordOut = new Object();

            promptStatus = myAuthPrompt.promptPassword(
                                            promptTitle,
                                            promptText,
                                            passwordRealm,
                                            Ci.nsIAuthPrompt.SAVE_PASSWORD_NEVER,
                                            passwordOut);

            // Canceled by user
            if(!promptStatus){
                this.logInfo("  - user canceled password prompt");
            } else {
                this.logInfo("  - password received from user");

                password = passwordOut.value;
                this.cachePassword(loginInfo, password);
            }
        }

        return password;
    },

    // Internal methods.

    // Save Password to session cache
    cachePassword: function(loginInfo, password){
        loginInfo.password = password;
        this.loginCache.push(loginInfo);
    },

    logInfo: function(aMsg, aDebugLevel)
    {
        var prefB = Cc["@mozilla.org/preferences-service;1"].getService(
						Ci.nsIPrefBranch);

        this.debug = this.globalFunctions.safeGetBoolPref(prefB, "extensions.1st-setup.authentication.debug", false, true);
        if (this.debug) {
            this.globalFunctions.LOG("mivExchangeLoginManager: "+aMsg);
        }
    },
}
