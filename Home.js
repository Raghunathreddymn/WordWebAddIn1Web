
(function () {
    //"use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
           
            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");
                
                $('#highlight-button').click(displaySelectedText);
                return;
            }
           
            $("#template-description").text("This sample highlights the longest word in the text you have selected in the document.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the longest word.");
            
            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(hightlightLongestWord);
            $('#Respond').click(test);
            $('#sharedoc').click(share);
            $('#btnreview').click(send); 
            $('#compare').click(compare); 
            
           
            
        });
    };

    function loadSampleData() {
        // Run a batch operation against the Word object model.
        //context.document.applicationName
        //Word.run(function (context) {
        //    // Create a proxy object for the document body.
        //    var body = context.document.body;
        //   // var body1 = context.document.properties.applicationName;

        //    // Queue a commmand to clear the contents of the body.
        //    body.clear();
        //    // Queue a command to insert text into the end of the Word document body.
        //    body.insertText(
        //       "context.document.properties.creationDate",


        //        Word.InsertLocation.end);
        //    context.sync().then

        //    // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
        //    return context.sync().then(function () {
        //        body.insertText(context.document.properties.creationDate, Word.InsertLocation.end)
        //    })
        //})
        //    .catch(errorHandler);

        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;

            // Queue a command to load content control properties.
          ///  context.load(thisDocument, 'contentControls/id, contentControls/text, contentControls/tag');
            thisDocument.properties.title = "test";
            console.log(thisDocument.properties.title);
           
           
            // Synchronize the document state by executing the queued-up commands, 
            // and return a promise to indicate task completion.
            return context.sync();
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        
    }

    function hightlightLongestWord() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();
            
            // This variable will keep the search results for the longest word.
            var searchResults;
            
            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
        .catch(errorHandler);
    } 


    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
    function test()
    {
        var ctx = new Word.RequestContext();
        ctx.sync();
        var author = ctx.document.properties.lastAuthor;
//        //showNotification("Error:", error);
//        window.open("Respond.html", "hello", `toolbar=no,directories=no,titlebar=no,scrollbars=no,resizable=no,status=no,location=no,toolbar=no,menubar=no,
//width=300,height=300,left=300,top=300`);
    }
    function share() {
        //showNotification("Error:", error);
        window.open("ShareDocument.html", "hello", `toolbar=no,directories=no,titlebar=no,scrollbars=no,resizable=no,status=no,location=no,toolbar=no,menubar=no,
width=500,height=500,left=300,top=300`);
    }

    function compare() {
      
    }
    function send() {
        console.log("before");
        
        var url = "https://btserviceapiappservice.azurewebsites.net/api/Configuration/sendforreview";

        var xhr = new XMLHttpRequest();
        xhr.open("POST", url);

        xhr.setRequestHeader("Accept", "application/json");
        xhr.setRequestHeader("Content-Type", "application/json");

        xhr.onreadystatechange = function () {
            if (xhr.readyState === 4) {
                console.log(xhr.status);
                console.log(xhr.responseText);
            }
        };
        let referenceid = CreateGuid();
        let emailids = document.getElementById("txtTo");
        let message = document.getElementById("txtMessage");
        let link = "https://blueed-my.sharepoint.com/:w:/g/personal/raghunath_mn_blueed_onmicrosoft_com/EYw8Z4FuHetEqom_atHoFZIBkYcjw5Eg0Hl5QJgs2TITTw?e=wWhBTe";
        var data = `{"documentReferenceID": ` + referenceid + `," emailIDs":` + emailids
            `,"message":` + message + `"link":` + link+`}`;
        document.getElementById("txtTo").innerHTML.value = Word.DocumentProperties.JSON;
       // xhr.send(data);


      
          
    }
    function CreateGuid() {
        function _p8(s) {
            var p = (Math.random().toString(16) + "000000000").substr(2, 8);
            return s ? "-" + p.substr(0, 4) + "-" + p.substr(4, 4) : p;
        }
        return _p8() + _p8(true) + _p8(true) + _p8();
    }

    
})();
