<!DOCTYPE html>
<!-- #include file = "secure.asp" -->
<%
if request("item")="56" then response.redirect "search_e.asp?year=2021"
if request("item")="57" then response.redirect "search_e.asp?year=2021&type=חותמות"
if request("item")="49" then response.redirect "store_e.asp"
sql = "select * from [stamps].[dbo].[pages] where theIndex='"&request("item")&"'"
r.open sql,strconn,1,3
if not r.eof then
    theType = r("pageType")
    thePage = r("theIndex")
    theTitle = r("theTitle")
    googleTitle = r("googleTitle")
    googleDesc = r("googleDesc")    
    theSubTitle = r("subTitle")
    theText = r("theText")
end if
r.close
if request("theType")="9" then
    theType=9
end if
%>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta charset="utf-8" />
    <link rel="stylesheet" href="css/styles_e.css" type="text/css" />
    <title><%=googleTitle%></title>
</head>
<body>
    <% 
    language="eng"
    %>
    <!-- #include file = "menu.asp" -->
    <div class="innerContainer">
        <!-- #include file="searchMenu_e.asp" -->
        <% 
        if request("theType")=9 then theTitle="News"
        %>
        <div class="innerTitle fullBox"><%=theTitle%></div>
        <div class="fullBox whiteBox">
            <div class="topLine" style="background-color:#7E237D"></div>
            <% 
            if theType=1 or theType=3 then
            %>
            <div style="float:right;width:100%;font-size:1.2em;padding:20px;box-sizing:border-box">
                <%=theSubTitle%>
            </div>
            <div style="float:right;width:60%;font-size:1em;padding:20px;box-sizing:border-box;text-align:justify">
                <%=theText%>
            </div>
            <div style="float:right;width:40%;box-sizing:border-box;padding:20px">
                <img src="cuts/stamp_big_gallery.jpg" style="width:100%" />
            </div>
            <%
            end if
            if request("theType")=9 then
                if request("item")<>"" then
                    sql = "select * from [stamps].[dbo].[News] where recId="&request("item")
                    r.open sql,strconn,1,3
                    if not r.eof then
                        %>
                        <div style="float:right;width:100%;font-size:1.2em;padding:20px;box-sizing:border-box;cursor:pointer;background-color:#DADADA">
                            <%=r("title_en")%><br /><%=r("post_date")%><br />
                            <% 
                            if trim(r("link"))&"a"<>"a" then
                                %>
                                <a href="<%=r("link")%>" target="_blank">Link</a>
                                <%
                            end if
                            %>
                        </div>
                        <div id="c<%=k%>" style="float:right;width:60%;font-size:1em;padding:20px;box-sizing:border-box;text-align:justify">
                            <%=r("content_en")%>
                        </div>
                        <div style="float:right;width:100%;font-size:1.2em;padding:20px;box-sizing:border-box;cursor:pointer;background-color:#DADADA">
                            <a href="info_e.asp?theType=9">All news & Events</a>
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
                            <%=r("title_en")%><br /><%=r("post_date")%><br />
                            <% 
                            if trim(r("link"))&"a"<>"a" then
                                %>
                                <a href="<%=r("link")%>" target="_blank">Link</a>
                                <%
                            end if
                            %>
                        </div>
                        <div id="c<%=k%>" style="float:right;width:60%;font-size:1em;padding:20px;box-sizing:border-box;text-align:justify">
                            <%=r("content_en")%>
                        </div>
                        <%
                        r.movenext
                    wend
                    r.close
                end if
            end if
            if theType=2 or theType=3 then
                sql = "select * from [stamps].[dbo].[list] where thePage = "&thePage&" order by theOrder"
                r.open sql,strconn,1,3
                i=0
                %>
                <div style="float:right;width:100%;box-sizing:border-box;padding:20px">
                    <div style="float:right;width:100%;font-size:1.2em;padding:20px;box-sizing:border-box;padding-top:5px">Items list</div>
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
                        <div style="float:left;width:100%;font-size:1.2em;background-color:<%=clr%>;box-sizing:border-box;padding-right:20px;height:40px;padding-top:5px">
                            <a href="images/<%=r("theFile")%>" target="_blank" style="width:80%"><%=r("theName")%></a>
                        </div>
                        <% 
                        elseif trim(r("theLink"))&"a"<>"a" then
                        %>
                        <div style="float:left;width:100%;font-size:1.2em;background-color:<%=clr%>;box-sizing:border-box;padding-right:20px;height:40px;padding-top:5px">
                            <a href="<%=r("theLink")%>" style="width:80%" target="_blank"><%=r("theName")%></a>
                        </div>
                        <% 
                        else
                        %>
                        <div style="float:left;width:100%;font-size:1.2em;background-color:<%=clr%>;box-sizing:border-box;padding-right:20px;height:40px;padding-top:5px">
                            <a href="info_e.asp?item=<%=request("item")%>&index=<%=r("theIndex")%>" style="width:80%"><%=r("theName")%></a>
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
            if theType=4 then
                sql = "select count(theIndex) as counter from [stamps].[dbo].[list] where thePage = "&thePage
                r.open sql,strconn,1,3
                lines = int(r("counter")/5)+1
                height = 130 * lines
                r.close
                sql = "select * from [stamps].[dbo].[list] where thePage = "&thePage&" order by theOrder"
                r.open sql,strconn,1,3
                i=0
                %>
                <div style="float:right;width:100%;box-sizing:border-box;padding:20px;height:<%=height%>">
                    <div style="float:right;width:100%;font-size:1.2em;padding:20px;box-sizing:border-box;padding-top:5px">Items list</div>
                    <%
                    while not r.eof
                        %>
                        <div style="float:left;width:19%;font-size:1.2em;background-color:#DADADA;box-sizing:border-box;padding-top:5px;margin-right:1%;height:220px;margin-bottom:10px;text-align:center">
                            <% 
                            if r("theImage")&"a"<>"a" then
                            %>
                            <div style="float:left;width:100%;height:67%"><img src="images/<%=r("theImage")%>" style="max-width:90%;max-height:98%" /></div>
                            <% 
                            else
                            %>
                            <div style="float:left;width:100%;height:67%">&nbsp;</div>
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
        <!-- #include file = "footer_e.asp" -->
</body>
</html>
