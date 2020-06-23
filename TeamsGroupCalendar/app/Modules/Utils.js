const getConfig = () => {
   var config = {
        clientId : "7fc9230c-fcd2-47fb-a2a3-6fb9a5ecda62",
        redirectUri : "/TeamsGroupCalendar/app/silent-end.html",
        authwindow :  "/TeamsGroupCalendar/app/auth.html",
        content: "/TeamGroupCalendar/app/tab.html",
        hostRoot: "https://atlasoutdoor.github.io",
   };
   return config;
}

const Authuser = (upn,appconfig) => {
    return new Promise(
        (resolve, reject) => {
            let config = {
                clientId: appConfig.clientId,
                redirectUri: window.location.origin + appConfig.redirectUri,       // This should be in the list of redirect uris for the AAD app
                cacheLocation: "localStorage",
                navigateToLoginRequestUrl: false,
                endpoints: {
                    "https://graph.microsoft.com": "https://graph.microsoft.com"
                }
            };
            if (upn) {
                config.extraQueryParameters = "scope=openid+profile&login_hint=" + encodeURIComponent(upn);
            } else {
                config.extraQueryParameters = "scope=openid+profile";
            }
            let authContext = new AuthenticationContext(config);
            console.log("successfully made authContext")
            let user = authContext.getCachedUser();
            console.log("got user", user);
            if (user) {
                if (user.userName !== upn) {
                    // User doesn't match, clear the cache
                    authContext.clearCache();
                }
            }
            // Get the id token (which is the access token for resource = clientId)
            let token = authContext.getCachedToken(config.clientId);
            console.log("token:", token);

            if (token) {
                authContext.acquireToken("https://graph.microsoft.com", function (error, idtoken) {
                    if (error || !idtoken) {
                       reject(error);
                    }
                    else
                        resolve(idtoken);
                });
            } else {
                console.log("no token! renewing...");
                // No token, or token is expired
                authContext._renewIdToken(function (err, idToken) {
                    if (err) {
                        console.log("Renewal failed: " + err);
                        microsoftTeams.authentication.authenticate({
                            url: window.location.origin + appConfig.authwindow,
                            width: 400,
                            height: 400,
                            successCallback: function (t) {
                                // Note: token is only good for one hour
                                token = t;
                                resolve(token);
                            },
                            failureCallback: function (err) {
                                  reject(err);
                            }
                        });
                    } else {
                        console.log("Token Renewal succeeded");
                        authContext.acquireToken("https://graph.microsoft.com", function (error, idtoken) {
                            if (error || !idtoken) {
                               reject(error);
                            }
                            else
                                resolve(idtoken);
                        });
                    }
                });
            }



        }
        );
}

const GetGroupMembers = (idToken, teamscontext) => {
    console.log("Getting group members with idToken: " + idToken);
    console.log("teamscontext: ", teamscontext);
    return new Promise(
        (resolve, reject) => {
            GroupId = teamscontext.groupId;
            if (!GroupId) {
                GroupId = "19a79608-753b-4ff8-81dc-f12783060fe0" // sales team
            }
            $.ajax({
                type: "GET",
                contentType: "application/json; charset=utf-8",
                url: ("https://graph.microsoft.com/v1.0/groups/" + GroupId + "/members"),
                dataType: 'json',
                headers: { 'Authorization': 'Bearer ' + idToken }
            }).done(function (item) {
                resolve(item);
            }).fail(function (error) {
                console.log("GetGroupMembers error: ", error);
                reject(error);
            });
        }
    );
}
const GetSchedule = (idToken, GroupMembers, displayNameMap) => {
    console.log("Getting schedule...");
    return new Promise(
        (resolve, reject) => {
            var SchPost = {};
            SchPost.schedules = [];
            for (index = 0; index < GroupMembers.value.length; ++index) {
                var entry = GroupMembers.value[index];
                SchPost.schedules.push(entry.mail);
                var dnMapValue = {};
                dnMapValue.displayName = entry.displayName;
                var initials = "";
                if(entry.givenName.length > 0){
                    initials = entry.givenName.slice(0,1);
                }
                if(entry.surname.length > 0){
                    initials = initials + entry.surname.slice(0,1);
                }
                dnMapValue.initials = initials;
                dnMapValue.colorEntry =  randomColor({luminosity: 'bright',format: 'hsla'});            
                displayNameMap[entry.mail] = dnMapValue;
            }
            var Start = new Date();
            var End = new Date();
            End.setDate(End.getDate() + 62); 
            var StartTime = {};
            StartTime.TimeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
            StartTime.dateTime = formatDate(Start) + "T08:00:00";
            var EndTime = {};
            EndTime.dateTime = formatDate(End) + "T08:00:00";
            EndTime.TimeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
            SchPost.startTime = StartTime;
            SchPost.endTime = EndTime;
            SchPost.availabilityViewInterval = 1440;
            var schRequest = JSON.stringify(SchPost);
            $.ajax({
                type: "POST",
                contentType: "application/json; charset=utf-8",
                url: ("https://graph.microsoft.com/beta/me/calendar/getSchedule"),
                dataType: 'json',
                data: schRequest,
                headers: {
                    'Authorization': 'Bearer ' + idToken,
                    'Prefer': ('outlook.timezone="' + Intl.DateTimeFormat().resolvedOptions().timeZone + "\""),

                }
            }).done(function (item) {
                resolve(item);
            }).fail(function (error) {
                reject(error);
            });
        }
    );
}

const GetUserPhotos = (idToken,GroupMembers) => {
    console.log("Getting user photos...");
    var photoRequestMap = {};
    for (index = 0; index < GroupMembers.value.length; ++index) {
        var entry = GroupMembers.value[index];
        var clientid = uuidv4();
        photoRequestMap[clientid] = entry.mail;
        var userImageURL = "https://graph.microsoft.com/v1.0/users('" + entry.mail + "')/photos/48x48/$value";
        var xhr = new XMLHttpRequest();
        xhr.open('GET', userImageURL, true);
        xhr.setRequestHeader("Authorization", "Bearer " + idToken);
        xhr.setRequestHeader("client-request-id", clientid);
        xhr.responseType = 'arraybuffer';
        xhr.onload = function (e) {
            if (this.status == 200) {
                // get binary data as a response
                var blob = this.response;
                var clientRespHeader = this.getResponseHeader("client-request-id")
                var ElemendId = "img" + photoRequestMap[clientRespHeader];
                var uInt8Array = new Uint8Array(this.response);
                var data = String.fromCharCode.apply(String, uInt8Array);
                var base64 = window.btoa(data);
                document.getElementById(ElemendId).src = "data:image/png;base64," + base64;

            } else {
                var clientRespHeader = this.getResponseHeader("client-request-id")
                var ElemendId = "img" + photoRequestMap[clientRespHeader];
                document.getElementById(ElemendId).src = "data:image/png;base64,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";
            }
        };

        xhr.send()
        $('#ShowBoard').hide();
        $('#ProgresLoader').hide();
    }
}

function formatDate(date) {
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();

    if (month.length < 2) month = '0' + month;
    if (day.length < 2) day = '0' + day;

    return [year, month, day].join('-');
}

function addDays(date, days) {
    var result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
}



const buildScheduleTable = (Schedules,displayNameMap) => {
    console.log("Building schedule table...");
    var JSONData = [];   
    for (index = 0; index < Schedules.value.length; ++index) {
        var entry = Schedules.value[index].scheduleItems;
        entry.forEach(function (CalendarEntry) { 
            console.log("CalendarEntry: ", CalendarEntry);
            calEntry ={};
            calEntry.title = CalendarEntry.subject + " (" + displayNameMap[Schedules.value[index].scheduleId].initials + ")";            
            if(CalendarEntry.start.dateTime.slice(12,8) == "00:00:00"){ 
                calEntry.start = CalendarEntry.start.dateTime.slice(0,11);
                calEntry.end = CalendarEntry.end.dateTime.slice(0,11);    
            }else{
                calEntry.start = CalendarEntry.start.dateTime;
                calEntry.end = CalendarEntry.end.dateTime;    
            }        
            calEntry.color =  displayNameMap[Schedules.value[index].scheduleId].colorEntry;
            JSONData.push(calEntry);
        });  

    }
    return JSONData;
}

const buildLegend = (displayNameMap) => {
    var html = "<div class=\"ms-Table\" style=\"border-collapse:collapse;border: 0px;table-layout: auto;width:100%;\;background-color:white;\"><div class=\"ms-Table-row\">";
    html = html + "<span class=\"ms-Table-cell\" style=\"background-color:white;font-size: large;width:50px;font-weight:bolder;\"></span>";
    html = html + "<span class=\"ms-Table-cell\" style=\"background-color:white;font-size: large;width:150px;font-weight:bolder;\">Member</span>";
    html = html + "</div>";
    for (var key in displayNameMap) {
        if (displayNameMap.hasOwnProperty(key)) {
            console.log(key);
            console.log(displayNameMap[key].colorEntry);
            html = html + "<div class=\"ms-Table-row\"><span class=\"ms-Table-cell\" style=\"width:50px;\"><img id=\"img" + key + "\" style=\"border: 2px solid " + displayNameMap[key].colorEntry  + ";\" src=\"\" /></span>";
            html = html + "<span class=\"ms-Table-cell ms-fontWeight-semibold\" style=\"vertical-align: middle;width:150px;background-color:" + displayNameMap[key].colorEntry + ";\">" + displayNameMap[key].displayName + "</span>";
            html = html + "</div >";
        }
    }
    html = html + "</div>";
    return html;
}

const uuidv4 = () => {
    function hex(s, b) {
        return s +
            (b >>> 4).toString(16) +  // high nibble
            (b & 0b1111).toString(16);   // low nibble
    }

    let r = crypto.getRandomValues(new Uint8Array(16));

    r[6] = r[6] >>> 4 | 0b01000000; // Set type 4: 0100
    r[8] = r[8] >>> 3 | 0b10000000; // Set variant: 100

    return r.slice(0, 4).reduce(hex, '') +
        r.slice(4, 6).reduce(hex, '-') +
        r.slice(6, 8).reduce(hex, '-') +
        r.slice(8, 10).reduce(hex, '-') +
        r.slice(10, 16).reduce(hex, '-');
}

