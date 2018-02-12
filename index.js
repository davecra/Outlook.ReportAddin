/// <reference path="node_modules/office-ui-fabric-js/dist/js/fabric.js" />
/// <reference path="node_modules/easyews/easyews.js" />

(function () {
    "use strict";

    var messageBanner;
    var securityEmailAddress;
    var headers;

    /**
     * The Office initialize function must be run each time a new page is loaded.
     * @param {string[]} reason 
     */
    Office.initialize = function (reason) {
        $(document).ready(function () {
            securityEmailAddress = getParam("email");
            loadForm();
        });
    };

    /**
     * Loads the form items, controls and attaches to events
     */
    function loadForm() {
        $("#forward-button").click(function () {
            getMessageHeaders();
        });

        $("#cancel-button").click(function() {
            Office.context.ui.closeContainer();
        });

        // banner
        messageBanner = new fabric["MessageBanner"](document.querySelector(".ms-MessageBanner"));
        $(".ms-MessageBanner").hide();

        // choice fields
        var ChoiceFieldGroupElements = document.querySelectorAll(".ms-ChoiceFieldGroup");
        for (var i = 0; i < ChoiceFieldGroupElements.length; i++) {
          new fabric['ChoiceFieldGroup'](ChoiceFieldGroupElements[i]);
        }

        // text fields
        var TextFieldElements = document.querySelectorAll(".ms-TextField");
        for (var i = 0; i < TextFieldElements.length; i++) {
          new fabric['TextField'](TextFieldElements[i]);
        }
    };

    /**
     * First step in response, we collect the headers then
     * we call to get the message info
     */
    function getMessageHeaders() {
        var itemId = Office.context.mailbox.item.itemId;
        var headers = easyEws.getEwsHeaders(itemId, getCurrentMessage, errorCallback);
    }

    /**
     * This function handles the click event of the sendNow button.
     * It retrieves the current mail item, so that we can get its itemId property
     * and also get the MIME content
     * It also retrieves the mailbox, so that we can make an EWS request
     * to get more properties of the item. 
     * @param {Dictionary} headersDictionary Contains all the messages headers from easyEws
     */
    function getCurrentMessage(headersDictionary) {
        headers = "\r\n\r\nInternet Headers:\r\n";
        headersDictionary.forEach(function(key, value) {
            headers += key + " = " + value + "\r\n";
        });
        var itemId = Office.context.mailbox.item.itemId;
        try {
            easyEws.getMailItemMimeContent(itemId, sendMessageCallback, showErrorCallback);
        } catch (error) {
            showNotification("Unspecified error.", error.Message);
        }
    };

    /**
     * Cleans a string containing invalid XML characters
     * BAsed on this: https://web.archive.org/web/20140228010526/http://validchar.com/d/xml10/xml10_namestart
     * @param {string} input The incoming string to be searched
     */
    function cleanString(input) {
        var output = "";
        var re = "@<>;?{}[]\\^`";
        var wi = ["(at)", "(", ")", ":", ".", "(", ")", "(", ")", "/", "*", "'"];
        debugger;
        for (var i=0; i<input.length; i++) {
            if (input.charCodeAt(i) <= 127) {
                var idx = re.indexOf(input.charAt(i));
                if(idx >= 0){
                    output += wi[idx];
                } else {
                    output += input.charAt(i);
                }
            }
        }
        return output;
    }
    
    /**
     * Gets specified parameter (name) from the URL
     * @param {string} name 
     * @returns {string} The value of the parameter
     */
    function getParam(name) {
        var url = window.location.href;
        name = name.replace(/[\[\]]/g, "\\$&");
        var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
            results = regex.exec(url);
        if (!results) return null;
        if (!results[2]) return '';
        return decodeURIComponent(results[2].replace(/\+/g, " "));
    };

    /**
     * This function is the callback for the getMailItemMimeContent method
     * in the getCurrentMessage function.
     * In brief, it first checks for an error repsonse, but if all 
     * is OK t:ItemId element.
     * @param {string} content BASE64 string of the MIME content of the message
     */
    function sendMessageCallback(content) {
        headers = cleanString(headers);
        var comment = $("#forward-comment").val();
        if (comment == null || comment == '') {
            comment = "[user provided no comment]";
        }
        var reason = "";
        var checked_reason = document.querySelector('input[name = "reasonfieldgroup"]:checked');
        if(checked_reason != null) { 
            reason = checked_reason.value; 
        } else {
            showNotification("Error", "You must select a reason.");
            return;
        }

        try {
            easyEws.sendPlainTextEmailWithAttachment("Security Message with Item Attachment",
                                                     "The reason from the user is this message " + 
                                                     "appears to be " + reason + ".\r\n" + 
                                                     "The comment from the user is: \r\n" + comment + 
                                                     headers,
                                                     securityEmailAddress,
                                                     "Suspicious Email",
                                                     content,
                                                     successCallback,
                                                     showErrorCallback);
        }
        catch (error) {
            showNotification("Unspecified error.", error.Message);
        }
    };

    /**
     * This function is the callback for the easyEws sendPlainTextEmailWithAttachment
     * @param {string} result - The result message for a successful callback
     */
    function successCallback(result) {
        showNotification("Success", result);
    };

    /**
     * This function will display errors that occur we use this as a callback
     * for errors in easyEws
     * @param {string} error - The error string 
     */
    function showErrorCallback(error) {
        showNotification("Error", error);
    };

    /**
     * Helper function for displaying notifications
     * @param {string} header Header of the message
     * @param {string} content Content of the message
     */
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
         $(".ms-MessageBanner").show();
        setTimeout(function() {
            $(".ms-MessageBanner").hide();
        },5000);
    };
})();