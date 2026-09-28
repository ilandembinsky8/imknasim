<!DOCTYPE html>
<style>
    p{
        margin:0px;padding:0px;
    }
</style>
<!-- #include file = "secure.asp" -->
<%
if request("item")="27" then response.redirect "search.asp?year=2022"
if request("item")="28" then response.redirect "search.asp?year=2022&type=חותמות"
if request("item")="54" then response.redirect "store.asp"
if request("theType")="" then
    sql = "select * from [stamps].[dbo].[pages] where theIndex='"&request("item")&"'"
    r.open sql,strconn,1,3
    if not r.eof then
        if r("parent")=0 and request("item")<>"26" then 
            theParent=r("theIndex")
            r.close
            sql = "select * from [stamps].[dbo].[pages] where parent="&theParent&" order by theOrder"
            r.open sql,strconn,1,3
            if r.eof then 
                response.redirect "main.asp"   
            else
                response.write theParent
                response.redirect "info.asp?item="&r("theIndex")
            end if
        end if
        theType = r("pageType")
        thePage = r("theIndex")
        theTitle = r("theTitle")
        googleTitle = r("googleTitle")
        googleDesc = r("googleDesc")    
        theSubTitle = r("subTitle")
        theText = r("theText")
        theFile = r("theFile")
        theStamp = r("theStamp")
    end if
    r.close
    if request("index")<>"" then
        sql = "select * from [stamps].[dbo].[list] where thePage="&request("item")&" and theIndex="&request("index")
        r.open sql,strconn,1,3
        if not r.eof then
            theTitle = "סיפורי בולים"
            theSubTitle = r("theName")
            theText = r("theText")
            theStamp = r("theImage")
        end if
        r.close
    end if
else
    theType=request("theType")
end if
%>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta charset="utf-8" />
    <link rel="stylesheet" href="css/styles.css" type="text/css" />
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
    <title><%=googleTitle%></title>
    <script>
        function showHide(x){
            alert(document.getElementById(y).style.display);
            y = "c"+x;
            if (document.getElementById(y).style.display = "inline")
                document.getElementById(y).style.display = "none";
            else
                document.getElementById(y).style.display = "inline";
        }
        function getProducts() {
            var userInfo = '<soap: Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap: Body><GetClient xmlns="http://tempuri.org/"></GetClient></soap: Body ></soap: Envelope > ';
            $.ajax({
                url: "getProducts.asp",
                type: "GET",
                contentType: "application/json; charset='utf-8'",
                data: userInfo,
                success: function (data) {
                    console.log(data);
                    var obj = JSON.parse(data);
                    s = "";
                    for (i = 0; i < obj.items.length; i++) {
                        s = s + "<div style='float:right;width:100%;height:250px;margin-bottom:20px'>";
                        s = s + "<div style='float:right;width:30%;margin-left:3%'><img src='images/"+obj.items[i].theImage+"' style='width:100%' /></div>"
                        s = s + "<div style='float:right;width:36%;margin-left:3%'>"
                        s = s + "<div style='float:right;width:100%;color:#00386a'><div style='float:right;font-size:24px'>"+obj.items[i].theName+"</div><div style='float:right;width:4px;height:20px;background-color:#00386a;margin-left:20px;margin-right:20px'>&nbsp;</div><div style='float:right;font-size:18px'>"+obj.items[i].description+"</div></div>"
                        s = s + "<div style='float:right;width:100%;color:#00386a;margin-top:25px'><div style='float:right;font-size:24px'>קוד: </div><div style='float:right;font-size:18px;margin-top:2px'>12345678ג</div></div>"
                        s = s + "</div>"
                        s = s + "<div style='float:right;width:14%;height:38px;background-image:url(cuts/stamp_price.png);background-size:100% auto;background-repeat:no-repeat;font-size:24px;text-align:center;padding-top:8px;box-sizing:border-box;color:#00386a'>" + obj.items[i].thePrice + "</div>"
                        s = s + "<div style='float:right;width:14%;height:38px;background-image:url(cuts/stamp_price.png);background-size:100% auto;background-repeat:no-repeat;font-size:24px;text-align:center;padding-top:8px;box-sizing:border-box;color:#00386a'><img src='cuts/shopping_cart_icon2.png'></div>"
                        s = s + "</div>"
                    }   
                    document.getElementById("products").innerHTML = s;
                },
                error: function (errMsg) {
                    alert("There is no connection");
                }
            });
        }
    </script>
</head>
<body onload="getSubject()">
    <% 
    language="he"
    %>
    <!-- #include file = "menu.asp" -->
    <div class="innerContainer">
        <!-- #include file="searchMenu1.asp" -->
        <% 
        if theType=9 then theTitle="חדשות"
        %>
        <div class="innerTitle fullBox"><%=theTitle%></div>
        <div class="fullBox whiteBox">
            <div class="topLine" style="background-color:#7E237D"></div>
            <% 
            if theType=7 then
                %>
                <div id="products" style="float:right;width:65%;margin-top:20px">
                
                </div>
                <script>
                    getProducts();
                </script>
                <%
            end if
            if theType=9 then
                if request("item")<>"" then
                    sql = "select * from [stamps].[dbo].[News] where recId="&request("item")
                    r.open sql,strconn,1,3
                    if not r.eof then
                        %>
                        <div style="float:right;width:100%;font-size:1.2em;padding:20px;box-sizing:border-box;cursor:pointer;background-color:#DADADA">
                            <%=r("title_he")%><br /><%=r("post_date")%><br />
                            <% 
                            if trim(r("link"))&"a"<>"a" then
                                %>
                                <a href="<%=r("link")%>" target="_blank">קישור</a>
                                <%
                            end if
                            %>
                        </div>
                        <div id="c<%=k%>" style="float:right;width:60%;font-size:1em;padding:20px;box-sizing:border-box;text-align:justify">
                            <%=r("content_he")%>
                        </div>
                        <% 
                        if r("theFile")&"a"<>"a" then
                        %>
                        <div id="f<%=k%>" style="float:right;width:60%;font-size:1em;padding:20px;box-sizing:border-box;text-align:justify">
                            <a href="images/<%=r("theFile")%>">הורד קובץ</a>
                        </div>
                        <%
                        end if
                        %>
                        <div style="float:right;width:100%;font-size:1.2em;padding:20px;box-sizing:border-box;cursor:pointer;background-color:#DADADA">
                            <a href="info.asp?theType=9">כל האירועים והחדשות</a>
                        </div>
                        <%
                    end if
                    r.close
                else
                    sql = "select * from [stamps].[dbo].[News] order by right(post_date,4)+right(post_date,6)+left(post_date,2) desc"
                    r.open sql,strconn,1,3
                    i=1
                    if i<10 then
                        k = "0"+i
                    else
                        k = i
                    end if
                    while not r.eof
                        %>
                        <div id="t<%=k%>" style="float:right;width:100%;font-size:1.2em;padding:20px;box-sizing:border-box;cursor:pointer;background-color:#DADADA">
                            <%=r("title_he")%><br /><%=r("post_date")%><br />
                            <% 
                            if trim(r("link"))&"a"<>"a" then
                                %>
                                <a href="<%=r("link")%>" target="_blank">קישור</a>
                                <%
                            end if
                            %>
                        </div>
                        <div id="c<%=k%>" style="float:right;width:60%;font-size:1em;padding:20px;box-sizing:border-box;text-align:justify">
                            <%=r("content_he")%>
                        </div>
                        <% 
                        if r("theFile")&"a"<>"a" then
                        %>
                        <div id="f<%=k%>" style="float:right;width:60%;font-size:1em;padding:20px;box-sizing:border-box;text-align:justify">
                            <a href="images/<%=r("theFile")%>">הורד קובץ</a>
                        </div>
                        <%
                        end if
                        r.movenext
                    wend
                    r.close
                end if
            end if
            if theType=1 or theType=3 then
            %>
            <div style="float:right;width:100%;font-size:1.2em;padding:20px;box-sizing:border-box">
                <%=theSubTitle%>
            </div>
            <% 
            if trim(theStamp)&"a"<>"a" and request("item")<>"36" then
            %>
            <div style="float:right;width:60%;font-size:1em;padding:20px;box-sizing:border-box;text-align:justify">
            <% 
            else
            %>
            <div style="float:right;width:100%;font-size:1em;padding:20px;box-sizing:border-box;text-align:justify">
            <%
            end if
            %>
            <%=theText%><br /><br />
            <% 
            if theFile&"a"<>"a" then
            %>
            <a href="images/<%=theFile%>"><%=theFile%></a><br /><br />
            <%
            end if 
            %>
            </div>
            <% 
            if trim(theStamp)&"a"<>"a" and request("item")<>"36" then
            %>
            <div style="float:right;width:40%;box-sizing:border-box;padding:20px">
                <%
                if theStamp&"a"<>"a" then 
                %>
                <img src="images/<%=theStamp%>" style="width:100%" />
                <% 
                else
                %>
                <img src="cuts/stamp_big_gallery.jpg" style="width:100%" />
                <% 
                end if
                %>
            </div>
            <%
            end if
            end if
            if theType=2 or theType=3 then
                if request("item")="11" then
                    sql = "select * from [stamps].[dbo].[lexicon] order by TITLE_HE"
                    r.open sql,strconn,1,3
                    %>
                    <div style="float:right;width:100%;box-sizing:border-box;padding:20px">
                    <div style="float:right;width:20%;font-weight:bold">כותרת עברית</div>
                    <div style="float:right;width:20%;font-weight:bold">כותרת בשפה זרה</div>
                    <div style="float:right;width:40%;font-weight:bold">תיאור</div>
                    <did style="float:right;width:20%;font-weight:bold">תמונה</did>
                    <%
                    i=0
                    while not r.eof
                        if i mod 2 = 0 then
                            clr = "#EAEAEA"
                        else
                            clr = "#FAFAFA"
                        end if
                        i=i+1
                        %>
                        <div style="float:right;width:100%;background-color:<%=clr%>;padding-top:10px;padding-bottom:10px">
                        <div style="float:right;width:19%;margin-left:1%"><%=r("title_he")%>&nbsp;</div>
                        <div style="float:right;width:19%;margin-left:1%;text-align:left"><%=r("title_en")%>&nbsp;</div>
                        <div style="float:right;width:39%;margin-left:1%"><%=trim(r("description"))%>&nbsp;</div>
                        <% 
                        if r("image")&"a"<>"a" then
                        %>
                        <div style="float:right;width:20%;background-color:#DADADA;height:200px;text-align:center"><img src="images/lexicon/<%=r("image")%>" style="max-width:90%;max-height:90%"/>&nbsp;</div>
                        <% 
                        else
                        %>
                        <div style="float:right;width:20%;height:200px">&nbsp;</div>
                        <% 
                        end if
                        %>
                        </div>
                        <%
                        r.movenext
                    wend
                    r.close
                    %>
                    </div>
                    <%
                else
                    sql = "select * from [stamps].[dbo].[list] where thePage = "&thePage&" order by theOrder"
                    r.open sql,strconn,1,3
                    i=0
                    %>
                    <div style="float:right;width:100%;box-sizing:border-box;padding:20px">
                        <div style="float:right;width:100%;font-size:1.2em;padding:20px;box-sizing:border-box;padding-top:5px">רשימת קישורים</div>
                        <%
                        while not r.eof
                            if i mod 2 = 0 then
                                clr = "#EAEAEA"
                            else
                                clr = "#FAFAFA"
                            end if
                            i=i+1
                            if trim(r("theFile"))&"a"<>"a" then
                            %>
                            <div style="float:right;width:100%;font-size:1.2em;background-color:<%=clr%>;box-sizing:border-box;padding-right:20px;height:40px;padding-top:5px">
                                <a href="images/<%=r("theFile")%>" target="_blank" style="width:80%"><%=r("theName")%></a>
                            </div>
                            <% 
                            elseif trim(r("theLink"))&"a"<>"a" then
                            %>
                            <div style="float:right;width:100%;font-size:1.2em;background-color:<%=clr%>;box-sizing:border-box;padding-right:20px;height:40px;padding-top:5px">
                                <a href="<%=r("theLink")%>" style="width:80%" target="_blank"><%=r("theName")%></a>
                            </div>
                            <% 
                            else
                            %>
                            <div style="float:right;width:100%;font-size:1.2em;background-color:<%=clr%>;box-sizing:border-box;padding-right:20px;height:40px;padding-top:5px">
                                <a href="info.asp?item=<%=request("item")%>&index=<%=r("theIndex")%>" style="width:80%"><%=r("theName")%></a>
                            </div>
                            <% 
                            end if
                            r.movenext
                        wend
                        r.close
                        %>
                    </div>
                    <%
                end if
            end if
            if theType=4 then
                sql = "select count(theIndex) as counter from [stamps].[dbo].[list] where thePage = "&thePage
                r.open sql,strconn,1,3
                lines = int(r("counter")/5)+1
                height = 130 * lines
                r.close
                sql = "select * from [stamps].[dbo].[pages] where theIndex = "&request("item")
                r.open sql,strconn,1,3
                if not r.eof then theText = r("theText")
                r.close
                sql = "select * from [stamps].[dbo].[list] where thePage = "&thePage&" order by theOrder"
                r.open sql,strconn,1,3
                i=0
                %>
                <div style="float:right;width:100%;box-sizing:border-box;padding:20px;height:<%=height%>">
                    <div style="float:right;width:100%;font-size:1.2em;padding:20px;box-sizing:border-box;padding-top:5px"><%=theText%></div>
                    <%
                    while not r.eof
                        %>
                        <div style="float:right;width:19%;font-size:1.2em;background-color:#DADADA;box-sizing:border-box;padding-top:5px;margin-left:1%;height:260px;margin-bottom:10px;text-align:center">
                            <% 
                            if r("theImage")&"a"<>"a" then
                            %>
                            <div style="float:right;width:100%;height:67%"><img src="images/<%=r("theImage")%>" style="max-width:90%;max-height:98%" /></div>
                            <% 
                            else
                            %>
                            <div style="float:right;width:100%;height:67%">&nbsp;</div>
                            <% 
                            end if
                            %>
                            <div style="float:right;width:100%">
                            <%
                            if trim(r("theLink"))&"a"<>"a" then
                            %>
                            <a href="<%=r("theLink")%>" target="_blank"><%=r("theName")%></a>
                            <%
                            elseif trim(r("theFile"))&"a"<>"a" then
                            %>
                            <a href="images/<%=r("theFile")%>" target="_blank"><%=r("theName")%></a>
                            <% 
                            else
                            %>
                            <%=r("theName")%>
                            <%
                            end if
                            %>
                            </div>
                        </div>
                        <% 
                        r.movenext
                    wend
                    r.close
                    %>
                </div>
                <%
            end if
            %>
            
        </div>
        <div style="float:right;width:100%;height:55px"></div>
        </div>
        <!-- #include file = "footer.asp" -->
</body>
</html>
