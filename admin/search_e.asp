<!DOCTYPE html>

<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
<head>
    <!-- #include file = "secure.asp" -->
    <% 
    language="eng"
    %>
    <meta charset="utf-8" />
    <link rel="stylesheet" href="css/styles_e.css" type="text/css" />
    <script src="https://code.jquery.com/jquery-3.1.1.min.js"></script>
    <title></title>
    <script>
        var SearchResults;
        var currentPage;
        var value;
        //var pages;
        var theValue=2;
        var theType=2019;

        function UrlExists(url) {
            var http = $.ajax({
                type: "HEAD",
                url: url,
                async: false
            })
            return http.status == 200;
        }

        function getParameterByName(name, url) {
            if (!url) url = window.location.href;
            name = name.replace(/[\[\]]/g, '\\$&');
            var regex = new RegExp('[?&]' + name + '(=([^&#]*)|&|#|$)'),
                results = regex.exec(url);
            if (!results) return null;
            if (!results[2]) return '';
            return decodeURIComponent(results[2].replace(/\+/g, ' '));
        }

        function getSubject() {
            var userInfo = '<soap: Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap: Body><GetClient xmlns="http://tempuri.org/"></GetClient></soap:></soap:> ';
            $.ajax({
                url: "getSubject.asp",
                type: "GET",
                contentType: "application/json; charset='utf-8'",
                data: userInfo,
                success: function (data) {
                    console.log(data);
                    var x = JSON.parse(data);
                    var x1 = document.getElementById("subject");
                    for (i = 0; i < x.items.length; i++) {
                        var option = document.createElement("option");
                        option.value = x.items[i].index;
                        option.text = x.items[i].title_en;
                        x1.add(option);
                    }
                },
                error: {

                }
            });
        }


    function showCancels(z) {
        if (searchResults.items.length == 0) {
            s = "לא התקבלו תוצאות";
            document.getElementById("results").innerHTML = s;
        }
        else {
            pages = Math.ceil(searchResults.items.length / 12.0);
            if (z < 1) z = 1;
            if (z > pages) z = pages;
            currentPage = z;
            s = "";
            theLimit = searchResults.items.length;
            if (searchResults.items.length > 12) theLimit = 12;
            start = (z - 1) * 12;
            finish = z * 12;
            if (finish > searchResults.items.length)
                finish = searchResults.items.length;
            if (SearchResults.items.length > 12) {
                width = 162;
                s = s + "<div style='float:left;width:100%;margin-bottom:20px;margin-top:20px;margin-left:20px'><div style='float:left;width:" + width + "px;height:40px;box-sizing:border-box;padding-top:8px;margin-left:10px'>";
                s = s + "<div style='float:left;width:40px;height:30px;text-align:center;cursor:pointer;background-color:#EAEAEA;margin-right:5px;box-sizing:border-box;padding-top:7px'  onclick='showResults(" + (currentPage - 1) + ")'><img src='cuts/arrow_b_right.png'></div>"
                s = s + "<div style='float:left;width:40px'><input id='paging1' style='width:100%;height:30px;box-sizing:border-box;border:1px solid #AAAAAA;outline:none;text-align:center' onkeypress='return runScript1(event)'></div>";
                s = s + "<div style='float:left;width:40px;height:30px;text-align:center;cursor:pointer;background-color:#EAEAEA;margin-left:5px;box-sizing:border-box;padding-top:7px' onclick='showResults(" + (currentPage + 1) + ")'><img src='cuts/arrow_b_left.png'></div>"
                s = s + "</div><div style='float:left;width:180px;margin-top:10px;padding-top:5px'> of <span id='pages1'></span></div></div>";
            }
            for (i = start; i < finish; i++) {
                s1 = "http://static.israelphilately.org.il/images/cancels/" + searchResults.items[i].image;
                s2 = "http://stamps.nez.co.il/cuts/stamp_gallery.jpg";
                s = s + "<div style='float:left;width:275px;height:275px;background-color:#FAFAFA;margin-left:12.5px;margin-left:12.5px;margin-top:12.5px;margin-bottom:12.5px;text-align:center;box-sizing:border-box;padding:10px'>";
                s = s + "<a href='cancel.asp?item=" + searchResults.items[i].index + "'><div style='float:righ;width:100%;height:60%'>";
                s = s + "<img src='"+s1+"' onerror='"+s2+"' style='max-height:100%;max-width:60%' />";
                s = s + "</div><div style='float:left;width:100%;height:40%;font-size:0.9em;box-sizing:border-box;padding-top:10px;text-align:left'>" + searchResults.items[i].title_e + "<br>date: "+searchResults.items[i].theDate+"</div></a>";
                s = s + "</div>";
            }
            if (SearchResults.items.length > 12) {
                width = 162;
                s = s + "<div style='float:left;width:100%;margin-bottom:20px;margin-top:20px;margin-left:20px'><div style='float:left;width:" + width + "px;height:40px;box-sizing:border-box;padding-top:8px;margin-left:10px'>";
                s = s + "<div style='float:left;width:40px;height:30px;text-align:center;cursor:pointer;background-color:#EAEAEA;margin-right:5px;box-sizing:border-box;padding-top:7px'  onclick='showResults(" + (currentPage - 1) + ")'><img src='cuts/arrow_b_right.png'></div>"
                s = s + "<div style='float:left;width:40px'><input id='paging1' style='width:100%;height:30px;box-sizing:border-box;border:1px solid #AAAAAA;outline:none;text-align:center' onkeypress='return runScript1(event)'></div>";
                s = s + "<div style='float:left;width:40px;height:30px;text-align:center;cursor:pointer;background-color:#EAEAEA;margin-left:5px;box-sizing:border-box;padding-top:7px' onclick='showResults(" + (currentPage + 1) + ")'><img src='cuts/arrow_b_left.png'></div>"
                s = s + "</div><div style='float:left;width:180px;margin-top:10px;padding-top:5px'> of <span id='pages1'></span></div></div>";
            }
            document.getElementById("results").innerHTML = s;
            document.getElementById("paging").value = currentPage;
            document.getElementById("pages").innerHTML = pages;
            document.getElementById("paging1").value = currentPage;
            document.getElementById("pages1").innerHTML = pages;
        }  
    }
    function getCancel(result, type) {
            var userInfo = '<soap: Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap: Body><GetClient xmlns="http://tempuri.org/"></GetClient></soap:></soap:> ';
            Url = "getCancel_e.asp?result=" + value + "&type="+theType+"";
            $.ajax({
            url: Url,
            type: "GET",
            contentType: "application/json; charset='utf-8'",
            data: userInfo,
                success: function (data) {
                    console.log(data);
                    searchResults = JSON.parse(data);
                    console.log(searchResults);
                    document.getElementById("phrase").innerHTML = searchResults.search;
                    showCancels(1);
                },
                error: function (errMsg) {
                    console.log("no Connection"+errMsg);
                }
        });
        return false;
        }

        
        
    function showResults(z) {
        if (SearchResults.items.length == 0) {
            s = "no results";
            document.getElementById("results").innerHTML = s;
        }
        else {
            console.log(SearchResults);
            pages = Math.ceil(SearchResults.items.length / 12.0);
            if (z < 1) z = 1;
            if (z > pages) z = pages;
            currentPage = z;
            s = "";
            theLimit = SearchResults.items.length;
            if (SearchResults.items.length > 12) theLimit = 12;
            start = (z - 1) * 12;
            finish = z * 12;
            if (finish > SearchResults.items.length)
                finish = SearchResults.items.length;
            if (SearchResults.items.length > 12) {
                width = 162;
                s = s + "<div style='float:left;width:100%;margin-bottom:20px;margin-top:20px;margin-left:20px'><div style='float:left;width:" + width + "px;height:40px;box-sizing:border-box;padding-top:8px;margin-left:10px'>";
                s = s + "<div style='float:left;width:40px;height:30px;text-align:center;cursor:pointer;background-color:#EAEAEA;margin-right:5px;box-sizing:border-box;padding-top:7px'  onclick='showCancels(" + (currentPage - 1) + ")'><img src='cuts/arrow_b_right.png'></div>"
                s = s + "<div style='float:left;width:40px'><input id='paging1' style='width:100%;height:30px;box-sizing:border-box;border:1px solid #AAAAAA;outline:none;text-align:center' onkeypress='return runScript1(event)'></div>";
                s = s + "<div style='float:left;width:40px;height:30px;text-align:center;cursor:pointer;background-color:#EAEAEA;margin-left:5px;box-sizing:border-box;padding-top:7px' onclick='showCancels(" + (currentPage + 1) + ")'><img src='cuts/arrow_b_left.png'></div>"
                s = s + "</div><div style='float:left;width:180px;margin-top:10px;padding-top:5px'> of <span id='pages1'></span></div></div>";
            }
            for (i = start; i < finish; i++) {
                s1 = "http://static.israelphilately.org.il/images/stamps/" + SearchResults.items[i].filename + ".jpg";
                s2 = "http://stamps.nez.co.il/cuts/stamp_gallery.jpg";
                s = s + "<div style='float:left;width:275px;height:275px;background-color:#FAFAFA;margin-left:12.5px;margin-left:12.5px;margin-top:12.5px;margin-bottom:12.5px;text-align:center;box-sizing:border-box;padding:10px'>";
                s = s + "<a href='stamp.asp?item=" + SearchResults.items[i].index + "'><div style='float:righ;width:100%;height:60%'>";
                s = s + "<img src='"+s1+"' onerror='"+s2+"' style='max-height:100%;max-width:60%' />";
                s = s + "</div><div style='float:left;width:100%;height:40%;font-size:0.9em;box-sizing:border-box;padding-top:10px;text-align:left'>" + SearchResults.items[i].title_e + "<br>date: "+SearchResults.items[i].theDate+"<br>Series: " + SearchResults.items[i].series_e + "</div></a>";
                s = s + "</div>";
            }
            if (SearchResults.items.length > 12) {
               width = 162;
                s = s + "<div style='float:left;width:100%;margin-bottom:20px;margin-top:20px;margin-left:20px'><div style='float:left;width:" + width + "px;height:40px;box-sizing:border-box;padding-top:8px;margin-left:10px'>";
                s = s + "<div style='float:left;width:40px;height:30px;text-align:center;cursor:pointer;background-color:#EAEAEA;margin-right:5px;box-sizing:border-box;padding-top:7px'  onclick='showCancels(" + (currentPage - 1) + ")'><img src='cuts/arrow_b_right.png'></div>"
                s = s + "<div style='float:left;width:40px'><input id='paging1' style='width:100%;height:30px;box-sizing:border-box;border:1px solid #AAAAAA;outline:none;text-align:center' onkeypress='return runScript1(event)'></div>";
                s = s + "<div style='float:left;width:40px;height:30px;text-align:center;cursor:pointer;background-color:#EAEAEA;margin-left:5px;box-sizing:border-box;padding-top:7px' onclick='showCancels(" + (currentPage + 1) + ")'><img src='cuts/arrow_b_left.png'></div>"
                s = s + "</div><div style='float:left;width:180px;margin-top:10px;padding-top:5px'> of <span id='pages1'></span></div></div>";
            }
            document.getElementById("results").innerHTML = s;
            document.getElementById("paging").value = currentPage;
            document.getElementById("pages").innerHTML = pages;
            document.getElementById("paging1").value = currentPage;
            document.getElementById("pages1").innerHTML = pages;
        }
        }
        function getResult(result, type) {
            var userInfo = '<soap: Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap: Body><GetClient xmlns="http://tempuri.org/"></GetClient></soap:></soap:> ';
            Url = "getResult_e.asp?result=" + value + "&type="+type+"";
            $.ajax({
            url: Url,
            type: "GET",
            contentType: "application/json; charset='utf-8'",
            data: userInfo,
                success: function (data) {
                    SearchResults = JSON.parse(data);
                    console.log(SearchResults);
                    document.getElementById("phrase").innerHTML = SearchResults.search;
                    showResults(1);
                },
                error: function (errMsg) {
                    console.log("no Connection"+errMsg);
                }
        });
        return false;
        }
    function doSearch(type) {
        if (type == "1") {
            value = document.getElementById("text").value;
            theType=1;
            theValue=value;
        }
        else{ 
            if (document.getElementById("subject").value != "0"){
                value = document.getElementById("subject").value;
                theType=2;
                theValue=value;
            }
            else{
                value = document.getElementById("year1").value;
                type=3;
                theType=3;
                theValue=value;
            }
        }
        getResult(value, type);
    }

    function runScript2(e) {
    if (e.keyCode == 13) {
        value = document.getElementById("text").value;
        doSearch(1);
        return false;
    }
}

function runScript1(e) {
    //See notes about 'which' and 'key'
    if (e.keyCode == 13) {
        location.href="search.asp";
        var tb = document.getElementById("paging1").value;
        showResults(tb);
        return false;
    }
}
        
function runScript(e) {
    //See notes about 'which' and 'key'
    if (e.keyCode == 13) {
        location.href="search.asp";
        var tb = document.getElementById("paging").value;
        showResults(tb);
        return false;
    }
}

    function checkIfFromAnotherPage() {
        item = getParameterByName("item");
        subject = getParameterByName("subject");
        year1 = getParameterByName("year");
        if (item != "" && item != null) {
            document.getElementById("text").value = item;
            doSearch(1);
        }
        if (subject != "" && subject != null) {
            document.getElementById("subject").value = subject;
            doSearch(2);
        }
        if (year1 != "" && year1 != null) {
            document.getElementById("year1").value = year1;
            doSearch(3);
        }
        getSubject();
    }
    function changeAndSearch(x,y){
        document.getElementById("text").value = "";
        if (y==2){
            document.getElementById("year1").value = "0";
            doSearch(2);
        }
        if (y==3){
            document.getElementById("subject").value = "0";
            doSearch(2);
        }
    }
    function getOne(x){
        if (x==1) getResult(theValue,theType);
        if (x==2) getCancel();
    }
    </script>
</head>
<body onload="checkIfFromAnotherPage();getSubject();">
    <!-- #include file="menu.asp" -->
    <div class="innerContainer">
        <div class="innerTitle fullBox">Catalog Search</div>
        <div class="fullBox whiteBox" style="height:100px">
            <div class="topLine" style="background-color:#F15B25"></div>
            <div class="r100">
                <div class="freeText" style="width:150px" >
                    <input id="text" type="text" class="search" onkeypress="return runScript2(event)" placeholder="Free text" />
                </div>
                <div class="magni" style="cursor:pointer" onclick="doSearch(1)"><img src="cuts/search_icon.png" /></div>
                <div class="searchSpace"></div>
                <div class="searchText">Search by</div>
                <div class="searchLine"></div>
                <div class="searchText">Subject: </div>
                <div class="searchSelect">
                    <select id="subject" class="search" onchange="changeAndSearch(this.value,2)">
                        <option value="0">Choose subject</option>
                    </select>
                </div>
                <div class="searchSpace"></div>
                <div class="searchText">Year: </div>
                <div class="searchSelect">
                    <select id="year1" class="search" onchange="changeAndSearch(this.value,3)">
                        <option value="0">Choose year</option>
                        <% 
                        i=year(date)
                        while i>=1948
                        %>
                        <option value="<%=i%>"><%=i%></option>
                        <% 
                        i=i-1
                        wend
                        %>
                    </select>
                </div>
                <div class="searchSpace"></div>
                <div class="searchText">Catalog: </div>
                <div class="searchSelect">
                    <select id="cat1" class="search" onchange="getOne(this.value)">
                        <option value="0">Choose catalog</option>
                        <option value="1">stamps</option>
                        <option value="2">cancels</option>
                    </select>
                </div>
                <div class="searchButton" onclick="doSearch(2)">Search</div>
            </div>
        </div>
        <div class="innerTitle fullBox">Search results - <span id="phrase"></span></div>
        <div class="fullBox whiteBox">
            <div class="topLine" style="background-color:#7E237D"></div>
            <div id="results" style="float:left;width:100%"></div>
        </div>
        <div style="float:right;width:100%;height:55px"></div>
        </div>
            
            </div>
<!-- #include file = "footer.asp" -->
</body>
</html>
