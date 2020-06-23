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
                document.getElementById(ElemendId).src = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMgAAADICAYAAACtWK6eAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAALEwAACxMBAJqcGAAAFTZJREFUeJztnWmQXcV1gL83I83TCCKhBSMhtAEJSwBXYsKqKhaDBEhg2Wy2STBgnMQ4ZnPYDJgQmzh2bBeuGIclLAZMMA5eZJAwhE0gNlEsYosJIKGNbTQGAUISmnn5cd7AaPTevNP39b19b7/zVZ1SldR693TfPre30+eAYRiGYRiGYRiGYRi5oBRagYgZDkwGpgJTgInA2KqMqf45AugAytU/O6r/d31V1lX/XA10Aauqf3YBy4DFwJKqfJB2hVoRM5Dm6QB2BnbrJ7sA4zPW4zXgGWBRP3kBMTAjIWYg7owG9gH2rcoeyAiQR9YCC4EFwIPAw0B3UI0KhhlIY9qBPYFDq/KpsOo0RQV4HLgDmAc8BvQE1cgoJGXgCOBG5ItbiVRWATcAh5PfUdDICe3AIcB1wNuE77xZy9vAtcCMalsYBgCTgIuRXaHQnTQvshS4CNl5M1qQErKemAv0Er5D5lV6gNuRUcXWqy1AGTgJeJbwna9osgj4Eh+f1RgRsRlwFnJWELqjFV1WAGcCnU5voKDEPmwOA/4OOA/YKsDzu5HT7sXIafdSPj4JX1X99zVsfGoOG5+uD0fOXvpO4LdE1gZ9J/RTq/+eNa8B/wJchehuFIh24GRgOdl8VXuRKcg1wOnAgUhHzopPAJ8GzkB2op4hu7XVUuBEoC31Whpe2A94knQ7xQbgAeDbyNbwFpnUzI0tkI2I7yAn6T2k2yaPA9MyqZmRiMnAL0mvA7yOjBBHkU+DaMQo4BhkhHmT9NrpZmx7OFe0A6cB7+H/Zb8FXA4cQFyHZ0OQKdmVyFrId7utBk7Bpl3B2QV4BP/Tp98AM5GOFDtDEVeTOfifhi0AdsquKkYf7cA3kV0fXy9zOXA+sHWG9cgb2wAXIlu5vtp1HXA2NppkxmRgPv5e4JPAcciX1BA6gOORnTlf7XwPYoBGinwef46E84GDiP8sqBlKiJvJAvy0eTdwZKY1aBE6gMvw85IeAQ7GDMOFErKdvRA/7+BSbMT2xgTgIZp/KS8idzzMMJJTAj4LvISfEXxcturHxzTk/KGZF/EO4jtkTnb+KCML73dp7t2sBPbOWPdo+CKyA9LMC7gBccMw0mEccBPNvaO1yAGmoaQEXEBzjf4qMmc2smEm4pPVzDs7B5v+NmQIcDXNNfRPgM2zVtzgT4Cf0ty7u4LWOJxNRBn4Fckb93XEOc8Iyyya8/W6BVsvbkInEoYmaaPehq018sQ4mn+fwzLXOqdsDtxLsobsweaueaUNcd1Jei/lbmyqTCfJjaMLOQk38s0hJI8tdjctPJKUST4MP434ZBnFYFuSB8q4jRZckwwBbiVZg92BREU3isVI4H9I9s5vIa77OINSIvlW7hWYD0+R6UBuZSZ99y2x1jyfZA10MS3SQJFTQqKhJOkDZwfQN1OOI1nDnBVCWSNVvkmyvnBsCGWzYBrJfKu+FkJZIxNOxb0/rEXys0TFBJJ55Z4SQlkjU5IYyUoicpXvQDIb2bTKqMd5uPeP+USyYZPkJuDFQTQ1QnIJ7v3k0iCaeuTzuFe6ZbbzjI0okWwL+KgQyvpgMu4BFubRmu7OY5Eg1H1ZcqcigapbjaHAXbj1mW4KGC2lHffQPE8jdwpiZzsksPb1wBMMHhFyNRLz9jrgBFojpOdI3N1S7qFgcbdc97i7iNu3agLidewjcc8TSKjVLKPHZ822uDs4FuYQcRfcIh72EK9X7g7I1/9DmjeMgbIWucE3KavKZMwhuLnKr6MAYU7bcY+Ve24QTdNlNLLZkHbKgQryMfo+kkUrNi7ErS0eJOdTrdNwq9BtxLdjdTTppheoJ68iAfFiog3x3nZph68G0VTBZNxSELxBXNdky8B/kL1h9Jce4J/J+VfUkfFIGgptG6wmp7tarslrYgqwMAoZ3kMaR3+ZQ1xJNo/Arf4/D6NmffbDrQI/CaNmKoxHcgKGNoqBMh/ZMo2Fy3Grf24cGtuBp9ArvoR4LuOPIp/G0Sf3E8+d7hG4JWVdSE6mmifj9tJmhFHTO2XyNa2qJ7eSk47igcNxq/sJQbTsxzDcrPqGMGqmQrNRBLOU81NqgxDcjL7eryIfsmC4bOu+Qzy7VkcTvtO7yAbiSdE8Hreo8sHuFA3H7RLUN8Ko6Z3RuG075kX+j8BfU4+ci77eKwm0o3eWg5IvEk9soysI39mTygUptEcIhgGvoK/3GVkrWAZec1DwiKwVTIkdyMZ9JC1ZQzzT3CPR13sFGX+gT3RQ7hHicSe5jvCdvFn5ru9GCUQJuQqgrffxWSr2nINisfgHTSAdr9ys5R3iuXdzKPp6LyKjD7WLUvOzUioDziF85/YlJ3lum1CUcEvwmsnHeq6DQjHd8/Bx2Skvcq/ntgnJDPT1/l3aykxCf4nlSeIZPbYjfKf2Kb3Ec+e9hN7dpwdHT19XF4Qvo+/0P6wqFQMHhFbAMyVg/9BKeKIC/EBZtg3ZYEqFdmAZOktdTiRBvapcT/ivvm/5d68tFJYOZCtXU+8lOAwMLiPIweiHp8uRHZ9Y2CW0Aimwa2gFPLIeuFJZdjJwYBpKXIfOQjcgW6KxUMLtpmRR5DWfjZQDJqI/xL3K98PL6IPA/cb3wwMzlvCdOS0Z7rGd8sDt6Oq9CuUSQDvFmoH+dpp36wxMLIdqtYjpxiHo+95olEcQWgM5RlmuC/i9smxRiNlAYqvbXOCPyrKqRDwaAxkCHKZ86K3IGsQoBjHtNIIs1n+tLHsYiv6vMZA9kLvXGn6hLFck3g2tQIrEWDdtH9wS+FSjQhoD0YbneQPxvYqN1aEVSJEY63YPsgjX0LBv+zSQucg2W2x0A++HViIF3kc8e2NjA5JGQ0PTBjIaxTBURatU0aggNyJj40WkbjGi7Yt7AlsMVqCRgWiDb/UgiU9i5dnQCqTA/4ZWIEXuRGf8JWDvwQo0MpB9lQo9jBwkxsoDoRVIgQdDK5AiXcBjyrKD9nFfBnKfslxRuSe0AikQY536c5+ynLaPb0IH8AG6o/tDkj6kQDxPeNcQX/Ky57bJI7PQtcUaBjkPGmwE2RldXNcKMsWKnZiiQt4YWoEMeEhZrhPYsd4/DmYguykf8CxxbhcO5Abi2MauEJex16MbGfU11HX992EgjyvLFZ3lwH+FVsIDc4CXQiuREdq+Wbev+zCQRcpyMfBdin928J3QCmTIM8py2r6+ESvRLXJSuZ2VY64h/CI7qdyaQnvkmeno2mWZ6w8PV/5whbjzdddiS9xzeOdB3kNu3bUS49C3j1Ng752UP6p1CouN4wnf4V3l9FRaIt+U0N+E/TOXHz5M+aOtskCvxc8I3+m1MielNigCT6Jro+m1/nO9RfoU5cMXOygaG6cguRnzzsvkIBVZQJYoy02t9Zf1DEQ7V9U+PEbeR9yl8/yReAP5MnaHViQg2vczqdZf1jOQscofXaosFyuvIx0wj+3wFuIC9EpoRQKjfTc1Q7E2ayBdynIx8xJyLUB7apsFSxAnvCJMAdPmLWW5mn2+noFoAxubgQgrkA7529CKIJHb90JyEhr6nVbtoADoPVf/0uVHW4Qz0XtB+5QPgYuIJye6L3ZH135Ol+K0uc+neKhAjGwH3EF2xnE/8OeZ1Kx4TEXXhk6n6W8qfzSWhJBpMR25jZiWYTwGfCaz2hQT7Wn6Gy4/+o7yRwe98G58xL7A1UiYnWaN4j3EXb3VfOCSMgpdu2ojMgL6OXSQBO0FZjjipfAj4Al07bwW2Y36MZJOe/PMtS42m6Hry2tq/ed62aJ60IcljeESUSjakHXcZGBEVUBGmtXAq8iWbW8A3WJhCLpcNb1IkqiNMAMxYqcpA6lnBOuVD+9QljOMUGjd2NfV+stmDcTJh94wAqD9iDsZSM3CTTzcMEKh/YjXHBSaHUFiS+FlxId2p7Vmnx9Sp/BqdIk4R1N8l/dtkBhgOwLbAlsD4xF/tBFIFqYy0labLOIKzAZkprAGed/diGPfSmT37CXE5egFip2xWOtXWDN0bj0D0TohOjl45YTtgJnIQdueyElrKzKkKpsh9+y3q1NuPXIr7z4kvd4DFCuLmLaPOl0f/zW6w5UvuPxoQEYDpwFPk57bR6tIF3AZ8EmnNxCOv0ZXr/+u9Z/rrUG0I0jeI5qMBb6HXJq5lITxj4yNGMPH141/D/xVWHUa0tTdpmYNJM9hZE5AksScjUwjDP9MBx4FriS/GXNrXqWtgZOBaF1/a150D8xwJEDateiTjxrJKQFfQXzL8uhyr+2jNft8PQNZrPzRKcpyWTEaSST6udCKtCDbI0l59gytyAC0BqLt80AxA8d1AgsIv4htdfkjsEODd5UVJfRXN/7U5Yc7lT9aIT+Xpv6T8J3DROR58nEVYjw6fXupc+Jeb4r1AfCaUom6uRUyZDrw5dBKGB+xE/Ct0Eqg37VcjqMvFqQcOt4jJeCHgXUwNuUMdN4YaaL9eNft64MZiDbvR2gDmQXsElgHY1PKwNcD66Dtm6kayO7KcmlhU6v8cjxhwxBp+2aiJFCfRL/ACRW8oROJkRt6UWpSX/aq+/bSZYyDjnVnIINZ9wtIwIBGlIC9VSr7Zxrmcp939g/03H2U5dYAf6j3j4MZyHpgofIhiZOxN0kedtCMwfmLQM+dpiz3KIO48zeaHy5QPuQAZTnf5NG1wdgYpwM4j+yvLDdoH29kIA8qH7IXYfyetgrwTMON8QGeuSV6L+OmDORhZBHTiDbgYKVCPhnRuIgRmBAbONOpH9KqP71IH69LIwPpRp+H8FBlOZ+YG3v+CRH5RtsXH0V8teqi2aO+Q/mwmdS/wpsWMd0RjxXNl9wnQ9EbyLxGBTQG0vBHqmwJ7Kcsaxhp8Wnk2oMGLwbyGPokkMcqyxlGWmj74JvIJa9B0RhIDzBX+dAjkSHOMEJQBj6rLDsXRVBwrZ/MLcpyo5HMqoYRgpnASGXZX/h8cBkJrKXxa5nj88ENeEqpk0lYyYp5Sn26UM50tCPIOiRWloaZSLRCw8iSycAMZdlfoYwW6eKKfLOyXBtwosPvGoYPTkK/pazty060IwHYNEPYCrKJ/G5TrGJI2pSRK+IaXV7BYWBwGUF6kESUGrbGtnyN7Pgi+hjLV5NiSruJiKFoLPVp0j9FtRGkGJImJeA5pR4bkI+3GtfrkMvQu57sRhgHRqO1OBRJX6FhLpLeIVVmoP9yLCDdUcRGkGJIWpQQh0OtHgelqMtGSj3joFSaB4dmIMWQtJjloMOTZOg4+SUHxRamqJgZSDEkDdqQTq/V4biU9KhJB7KVq1VO6x/jihlIMSQNjnZ4/jIC+Aie6aDgS6RzccYMpBjim04kGrv2+aemoINKyZUOSp6dgg5mIMUQ35zv8OxlwLAUdFDxDwoF+2Q1/hNmmoEUQ3wyAbdggX/v+flOlNG7n1SAmzw/3wykGOKTXzo8dzFNujw1e4d8HXARcI2y/BeAnwO3N/ncPp6jWCmJjeaYDRzlUP5bSADEoLQhkU+0Vr2U/CZ8NPLLSNx2Th8hbODsjZiG25D70zBqGgXmKtz6WN5yJXIzbhWYFUZNo4DMxq1vXR9GzcGZCLyLvhJv4n9Xy4iPrZErstp+9TaOHrtZ8jXcLH0eOZonGrmjDbgLtz71t0E0VdKGeyrm84NoahSBf8KtL91PAT64OyPbv9pK9WKhgoxNcfHUrSDJnvKSn70hZ+NWuW5gahBNjTyyPfowU33yjSCaJqQNuAe3Cj6LPuiXES+jgOdx6zt3UoCp1UAmIiODS0XvwkKXtjIdwL249Zkucrxr1YijcKtsBYk6kXXIfCM8JeT8wrW/zA6hrE8uxb3SlwTR1AhFCfg+7v3kByGU9c1QYD7ulT8vhLJGEC7EvX/cS/YJm1JjHG6Xq/rk6yGUNTLldNz7xXLgEyGUTZO9kb1qMxKjjyTG8QE5dET0xbG4N4hNt+KjRLJpVS+SpClqziWZkVyC7W7FQNIFeYWCHQYmpQRcSbIGuho7JykyHSTbyq0Al9FCH8ghSFq3JA11F3biXkRG4X4I2Cc30YIpvzuA20jWYM8C22avspGQ7XF3H+mT39LCs4ZO4G6SNVw35gVcBGbh7njYf7YQLKZVXtic5EbSC1xAAR3VWoA23O9zDDSOzbJWOq8MI/l0q4LkLBmfudZGPbbG/SbgwGlVy48cA+kg+cK9ArwFHJ651sZAZuN2h3yg3EQLrzkaMQS4guSNWwEuB0ZkrbjBSNxD8wyUy2jB3SpXSsA5NNfQy7HRJEtm4xbUbaD0ItkCWuacwwfHksx3q7/cjK1N0mQCbrFya8kHtID7SFrsQzIv4P7yLuLeYos+f3Qiu4cuUdbrjfTROh5mxXjgAZp7ERXgZeBz2DDeDG3AMcASmn8f9wJbZap9xAwl2c3EWvI4kjrYDEVPCTnwc8kJOJj8GxFddsoTR+EeCKKePISktDZDqU8JOAyJlO6jzbuAz2RagxZkG9xDCg0mi5CsvU0lWomMMnAikn/FVzvfSYGjjxSNdiQ4nUsEx0ayAkkGNDHDeuSNScDFNL8x0l/WIlu45g4UgJ2AB/H3MitAD5IBazatMaqUkc2Lech5hM+2vJ8ChQONlTbgq0iSUJ8vt4Ksd64GphPXonIo4gl9Lck9bQeTt4GvYKNGrtgGyX/o+2X3SRdwA5JncUxGdfLJWOA44EZgFem10/XYWiPX7AMsJL0O0DcNexj4V2T7c3QmNXNjDOJu8z3gUfxPnwbKI9ihX2FoA04AXiXdTtFfngN+BvwjMiUbRzZbyCXkMHVG9dnXk/z2XhJZDPwNkU6nYj8DKAMnI0l6QvhkvYN0oCXVP5cibvmrkCnbKsQXaR2Srnhd9f+VkQ2CMuLWMQaZIvXJJCRdRJ+E8F5eAXwbWccET7VsNEcncAbNeZyaiCwDTsV826KkAzgeORwM3dGKJk8hi3y7zNQClJB1wu+QBXfozpdX2QDMAQ4i/um4UYdtkDCYSwjfIfMii5F1m23XGh/RhnwpryLdc4K8yiok4uWBRLojZfhjKOISfx2y6xS686YlbyK7UIdga4ua2LyyMW3A7ojBHArsQXHbrRc5NJxXlSeqf2fUoagvOiRbILlO9q3Knsg2ch5ZgxjEgqo8jJzNGErMQJpnKLAjsCuwW1V2RRb/WVFB7nc/U5VFVfkD8GGGekSHGUh6lIHJwBTktHsSm56Ij0TOZvpOzcvV/7uOj0/X1yPesX2n732yDNltWoyc0PedwhuGYRiGYRiGYRiGYcTH/wOJQIVW6eyMdwAAAABJRU5ErkJggg==";
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

