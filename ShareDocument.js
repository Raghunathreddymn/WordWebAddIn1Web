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
        `,"message":` + message + `"link":` + link + `}`;

    xhr.send(data);




}
    