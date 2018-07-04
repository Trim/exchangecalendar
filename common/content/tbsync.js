"use strict";

var ews = {
    bundle: Services.strings.createBundle("chrome://ews4tbsync/locale/ews.strings"),
    

    init: Task.async (function* ()  {
        //for example load overlays or do other init stuff
        tbSync.window.alert("EWS Test");	    
    }),

    //this is called after standard init, if lighning is available
    init4lightning: Task.async (function* () {
    }),

    //this is  called during tbSync cleanup (if lighning is enabled)
    cleanup4lightning: function ()  {
    },

    //CORE SYNC LOOP FUNCTION (syncdata contains account, folder etc.
    start: Task.async (function* (syncdata, job, folderID = "")  {
        //set syncdata for this sync process (reference from outer object)
        syncdata.syncstate = "";
        syncdata.folderID = folderID;
    }),

    getNewAccountEntry: function () {
        let row = {
            "account" : "",
            "accountname": "",
            "provider": "ews",
            "lastsynctime" : "0", 
            "status" : "disabled", //global status: disabled, OK, syncing, notsyncronized, nolightning
            "servertype" : "", //autodiscover or manual
            "host" : "",
            "user" : "",
            "https" : "1",
            "autosync" : "0"}; 
        return row;
    },
    
    getNewFolderEntry: function () {
        let folder = {
            "account" : "",
            "folderID" : "",
            "name" : "",
            "type" : "",
            "target" : "",
            "targetName" : "",
            "targetColor" : "",
            "selected" : "",
            "lastsynctime" : "",
            "status" : "",
            "parentID" : "",
            "cached" : "0"};
        return folder;
    },
    
	abServerSearch: Task.async (function* (account, currentQuery)  {
        for (let i=0; i < 5; i++) {
            let result = {Properties: {FirstName: currentQuery + "#" + i, LastName: "Test", DisplayName: "DisplayTest", EmailAddress: "user@inter.net"}};
            results.push(result);
        }
        return results;
    }),





    
    //used by accountSettings UI, not required by the manager itself (you can use any other implementation inside accountSettings.js)
    getAccountStorageFields: function () {
        return Object.keys(this.getNewAccountEntry()).sort();
    },

    //what settings should be locked due to autodiscover?
    getFixedServerSettings: function(servertype) {
        let settings = {};

        switch (servertype) {
            case "auto":
                settings["host"] = null;
                settings["https"] = null;
                break;
        }
        
        return settings;
    },
    
};
