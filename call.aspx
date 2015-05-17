<!DOCTYPE HTML PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%@ Page Language="C#" %>
<%@ Register tagprefix="SharePoint" namespace="Microsoft.SharePoint.WebControls" assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<html dir="ltr" xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
    <meta name="WebPartPageExpansion" content="full" />
    <meta http-equiv="X-UA-Compatible" content="IE=10" />
    <!--このページのSharePoint Onlineの配置先のサイトコレクションのURL/_layputs/15/defaultcss.ashxを参照してください。 Ref deploy sitecollection css -->
    <link rel="stylesheet" href="https://movaritest04.sharepoint.com/sites/portal/_layouts/15/defaultcss.ashx" />
</head>
<body>
    <div id="result"></div>
    <script>
        (function () {
            //Azure Active Directoryに登録したClientIDを指定してください
            var CLIENTID = "600cb9cc-d5ad-4a98-b755-0e0e5df48a8a";
            //SharePoint Onlineにこのページを保存したときのURLを指定してください
            var REDIRECTURL = "https://movaritest04.sharepoint.com/sites/portal/SitePages/call.aspx";
            //SharePoint OnlineのサイトコレクションのURLを指定してください
            var CURRENTSITEURL = "https://movaritest04.sharepoint.com/sites/portal";
            //Office365APIのURLのため、以下2項目は書き換える必要はありません(Preview版のOffice 365 unified APIを使っているので注意してください)
            var RESOURCEURL = "https://graph.microsoft.com";
            var AUTHENDPOINTURL = "https://login.microsoftonline.com/common/oauth2/authorize";

            //Office API利用に必要なAccess Tokenを取得後、検証処理を開始
            function initAccessToken() {
                if (location.hash) {
                    var hasharr = location.hash.substr(1).split("&");
                    hasharr.forEach(function (hashelem) {
                        var elemarr = hashelem.split("=");
                        if (elemarr[0] == "access_token") {
                            getCurrentSPUserAccountName(elemarr[1]);
                        }
                    }, this);
                } else {
                    location.href = AUTHENDPOINTURL + "?response_type=token&client_id=" + CLIENTID + "&resource=" + encodeURIComponent(RESOURCEURL) + "&redirect_uri=" + REDIRECTURL;
                }
            }

            //SharePoint OnlineのREST APIで現在のユーザーの識別情報を取得し、Access Tokenの検証後、予定の取得処理を開始
            function getCurrentSPUserAccountName(accesstoken) {
                var xhr = new XMLHttpRequest();
                xhr.open("GET", CURRENTSITEURL + "/_api/SP.UserProfiles.PeopleManager/GetMyProperties?$select=AccountName");
                xhr.setRequestHeader("Accept", "application/json;odata=verbose");
                xhr.onload = function () {
                    if (xhr.status == 200) {
                        var jsonFormattedResponse = JSON.parse(xhr.response);
                        var accountName = jsonFormattedResponse.d.AccountName.split("|")[2];
                        var jsonFormattedAccesstoken = JSON.parse(base64Util.decodeRFC4648(accesstoken.split(".")[1]) + '}');
                        var accessTokenUpn = jsonFormattedAccesstoken.upn;
                        if (accessTokenUpn === accountName) {
                            initMyCalendar(accesstoken);
                        }
                    } else {
                        document.getElementById("result").textContent =
                          "HTTP " + xhr.status + "<br>" + xhr.response;
                    }
                }
                xhr.send();
            }

            //Access Tokenに基づいてExchange Onlineの情報を取得
            function initMyCalendar(token) {
                var now = new Date();
                //日本時間の「今日」の予定を取得
                var startDt = [now.getFullYear(), (now.getMonth() + 1), (now.getDate() - 1)].join("-") + "T15:00:00Z";
                var endDt = [now.getFullYear(), (now.getMonth() + 1), (now.getDate())].join("-") + "T15:00:00Z";
                var xhr = new XMLHttpRequest();
                xhr.open("GET", "https://graph.microsoft.com/beta/me/calendarview?startDateTime=" + startDt + "&endDateTime=" + endDt + "&$select=Subject,Start,End,WebLink");
                console.log(token);
                xhr.setRequestHeader("Authorization", "Bearer " + token);
                xhr.onload = function () {
                    if (xhr.status == 200) {
                        var jsonFormattedResponse = JSON.parse(xhr.response);
                        var formattedResponse = "<table width='200px'><tr><th width='50%'>件名</th><th width='25%'>開始</th><th width='25%'>終了</th></tr>";
                        for (var i = 0; i < jsonFormattedResponse.value.length; i += 1) {
                            var start = utcDateUtil.parse(jsonFormattedResponse.value[i].Start, 9);
                            var end = utcDateUtil.parse(jsonFormattedResponse.value[i].End, 9);
                            formattedResponse += "<tr><td><a href='#' onclick='(function (){window.open(\""
                            		 + jsonFormattedResponse.value[i].WebLink
                            		 + "\", \"" + jsonFormattedResponse.value[i].Subject + "\", \"width=500,height=500,scrollbars=yes\")}())' >"
                            		 	 + jsonFormattedResponse.value[i].Subject
                                     + "</a></td><td>" + start.Hour + ":" + start.Minute
                                     + "</td><td>" + end.Hour + ":" + end.Minute + "</td></tr>";
                        }
                        formattedResponse += "</table>";
                        document.getElementById("result").innerHTML = formattedResponse;
                    } else {
                        document.getElementById("result").textContent =
                          "HTTP " + xhr.status + "<br>" + xhr.response;
                    }
                }
                xhr.send();
            }

            var utcDateUtil = {
                //SharePoint Onlineから取得した日時の情報をするための文字列処理
                "parse": function (utcDateString, timeDifference) {
                    var parsed = {};
                    var datearr = utcDateString.split("T")[0].split("-");
                    var timearr = utcDateString.split("T")[1].slice(0, -1).split(":");
                    parsed["Year"] = datearr[0];
                    parsed["Month"] = datearr[1];
                    parsed["Day"] = datearr[2] + Math.floor((Number(timearr[0]) + timeDifference) / 24);
                    parsed["Hour"] = (Number(timearr[0]) + timeDifference) % 24;
                    parsed["Minute"] = timearr[1];
                    parsed["Second"] = timearr[2];
                    return parsed;
                }
            }

            var base64Util = {
                //Access Tokenの情報をデコードするための文字列処理
                //Thanks for http://www.simplycalc.com/base64-source.php
                "decodeRFC4648": function (data) {
                    var b64pad = '=';
                    var dst = "";
                    var i, a, b, c, d, z;

                    function base64_charIndex(c) {
                        if (c == "+") return 62
                        if (c == "/") return 63
                        return b64u.indexOf(c)
                    }
                    
                    for (i = 0; i < data.length - 3; i += 4) {
                        a = base64_charIndex(data.charAt(i + 0));
                        b = base64_charIndex(data.charAt(i + 1));
                        c = base64_charIndex(data.charAt(i + 2));
                        d = base64_charIndex(data.charAt(i + 3));
                        dst += String.fromCharCode((a << 2) | (b >>> 4))
                        if (data.charAt(i + 2) != b64pad)
                            dst += String.fromCharCode(((b << 4) & 0xF0) | ((c >>> 2) & 0x0F));
                        if (data.charAt(i + 3) != b64pad)
                            dst += String.fromCharCode(((c << 6) & 0xC0) | d);
                    }
                    dst = decodeURIComponent(escape(dst));
                    return dst;
                }
            }

            //アクセストークンの取得処理を開始
            initAccessToken();

        }());
    </script>
</body>
</html>