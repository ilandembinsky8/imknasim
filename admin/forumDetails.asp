<!-- #include file="secure.inc" -->
 <script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
<script>
    function gotoPage(x){
        location.href="forumDetails.asp?num=<%=request("num")%>&item="+x+""
    }
    function deleteItem(x,y){
        //alert("deleteItem.asp?table="+x+"&item="+y+"&num=<%=request("num")%>");
        location.href="deleteItem.asp?table="+x+"&item="+y+"&num=<%=request("num")%>";
    }
    function editItem(x,y){
        var userInfo = '<soap: Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap: Body><GetClient xmlns="http://tempuri.org/"></GetClient></soap: Body ></soap: Envelope > ';
		    var url = "../server/getConfItem.asp?table="+x+"&item="+y+"";	
            $.ajax({
                url: url,
                type: "GET",
                contentType: "application/json; charset='utf-8'",
                data: userInfo,
			    success: function (data2) {
                    x1 = JSON.parse(data2);
                    console.log(x1);
                    if (x=="split" && y!=0){
                        document.getElementById("theNumber").value=x1.theNumber;
                        document.getElementById("theName").value=x1.theName;
                        document.getElementById("location").value=x1.location;
                        document.getElementById("theDay").value=parseInt(x1.theDay);
                        document.getElementById("theOrder").value=parseInt(x1.theOrder)
                    }
                    if (x=="session" && y!=0){
                        document.getElementById("theNumber").value=x1.theNumber;
                        document.getElementById("theName").value=x1.theName;
                        document.getElementById("split").value=x1.split;
                        document.getElementById("showinagenda").value=x1.showinagenda;
                        document.getElementById("theOrder").value=parseInt(x1.theOrder)
                        document.getElementById("theDay").value=parseInt(x1.theDay)
                        document.getElementById("chairName").value=x1.lecturers[0].theName;
                        document.getElementById("chairTitle").value=parseInt(x1.lecturers[0].theTitle)
                        document.getElementById("chairJob").value=x1.lecturers[0].theJob
                        document.getElementById("chairOrganization").value=parseInt(x1.lecturers[0].theOrganization);
                    }
                    if (x=="consponser" && y!=0){
                        document.getElementById("theNumber").value=x1.theNumber;
                        document.getElementById("sponsor").value=parseInt(x1.sponsor);
                        document.getElementById("theOrder").value=parseInt(x1.theOrder)
                    }
                    if (x=="item" && y!=0){
                        document.getElementById("theNumber").value=x1.theNumber;
                        document.getElementById("theName").value=x1.theName;
                        document.getElementById("session").value=parseInt(x1.session);
                        document.getElementById("sponsor").value=parseInt(x1.sponsor);
                        document.getElementById("itemType").value=parseInt(x1.itemType);
                        document.getElementById("fromHour").value = x1.fromHour;
                        document.getElementById("toHour").value = x1.toHour;
                        document.getElementById("theOrder").value=parseInt(x1.theOrder)
                        document.getElementById("lecName").value=x1.lecturers[0].theName;
                        document.getElementById("lecTitle").value=parseInt(x1.lecturers[0].theTitle)
                        document.getElementById("lecJob").value=x1.lecturers[0].theJob
                        document.getElementById("organization").value=parseInt(x1.lecturers[0].theOrganization);
                    }
                },
                error: function (xhr, ajaxOptions, thrownError) {
                    console.log("error");
                    return;
                }
		    });
        document.getElementById(x+"Edit").style.display="inline";
    }
    function doSubmit(x){
        const form = document.getElementById(x);
        //alert(x);
        // Your processing here
        form.submit();
    }
</script>
<%
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")


%>
<html><head><title>מערכת ניהול</title>
<meta charset="UTF-8">

</head>
<body style="margin:0px;padding:0px;direction:rtl;text-align:right">
<link rel="stylesheet" type="text/css" href="admin.css">
<div style="float:right;width:100%">
<div style="float:right;width:100%;text-align:center;font-size:40px;background-color:#FAFAFA;height:50px">מערכת ניהול</div>
<div style="float:right;width:100%;height:100%">
    <div style="float:right;width:20%;height:100%;background-color:#EAEAEA">
        <!-- #include file="toolbar.inc" -->
    </div>
    <div style="float:right;width:75%;box-sizing:border-box;padding:10px">
        <div style="float:right;width:100%;height:40px;font-size:2rem">
            <%
            sql = "select * from [imknasim].[dbo].[conference] where theindex="&request("num")&""
            r.open sql,strconn,1,3
            if not r.eof then 
                response.write r("theName")
            end if
            r.close
            %>
        </div>
        <div style="float:right;width:100%;height:40px">
            <div style="float:right;width:250px;height:100%;margin-left:20px;background-color:#DADADA;color:black;font-size:1.3rem;text-align:center;padding-top:8px;border-radius:10px;cursor:pointer;<%if request("item")="split" then response.write "font-weight:bold"%>" onclick="gotoPage('split')">פיצולים</div>
            <div style="float:right;width:250px;height:100%;margin-left:20px;background-color:#DADADA;color:black;font-size:1.3rem;text-align:center;padding-top:8px;border-radius:10px;cursor:pointer;<%if request("item")="session" then response.write "font-weight:bold"%>" onclick="gotoPage('session')">מושבים</div>
            <div style="float:right;width:250px;height:100%;margin-left:20px;background-color:#DADADA;color:black;font-size:1.3rem;text-align:center;padding-top:8px;border-radius:10px;cursor:pointer;<%if request("item")="consponser" then response.write "font-weight:bold"%>" onclick="gotoPage('consponser')">נותני חסות</div>
            <div style="float:right;width:250px;height:100%;margin-left:20px;background-color:#DADADA;color:black;font-size:1.3rem;text-align:center;padding-top:8px;border-radius:10px;cursor:pointer;<%if request("item")="item" then response.write "font-weight:bold"%>" onclick="gotoPage('item')">פריטים</div>
        </div>
        <div style="float:right;width:100%;height:auto;margin-top:20px">
            <%
            if request("item")="split" then
                sql = "select * from [imknasim].[dbo].[split] where theconference="&request("num")&" order by theOrder"
                r.open sql,strconn,1,3
                %>
                <div style="float:right;width:45%">
                <table style="font-size:1.5rem;padding:10px;margin:10px">
                <%
                if not r.eof then
                    %>
                    
                    <tr style="border-bottom:1px solid #DADADA;height:50px"><td>שם</td><td>מיקום</td><td>יום</td><td>סדר</td></tr>
                    <%
                    while not r.eof
                        %>
                        <tr style="border-bottom:1px solid #DADADA;height:50px"><td><%=r("theName")%></td><td><%=r("location")%></td><td><%=r("theday")%></td><td><%=r("theOrder")%></td>
                            <td style="cursor:pointer" onclick="editItem('<%=request("item")%>',<%=r("theIndex")%>)">ערוך</td>
                            <td style="cursor:pointer" onclick="deleteItem('<%=request("item")%>',<%=r("theIndex")%>)">מחק</td>
                        </tr>
                        <%
                        r.movenext 
                    wend
                end if
                %>
                <tr><td style="cursor:pointer" colspan=6 onclick="editItem('<%=request("item")%>',0)">הוסף</td></tr>
                </table>
                </div>
                <div id="splitEdit" style="float:right;width:50%;margin:10px;padding:10px;display:none;font-size:1.5rem">
                    <form id="splitForm" method="post" action="saveTable.asp">
                    <input type="hidden" id="theConf" name="theConf" value="<%=request("num")%>">
                    <input type="hidden" id="theTable" name="theTable" value="split">
                    <input type="hidden" id="theNumber" name="theNumber">
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">שם</div>
                        <div style="float:right;width:50%"><input type="text" name="theName" id="theName" style="width:100%"></div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">מיקום</div>
                        <div style="float:right;width:50%"><input type="text" name="location" id="location" style="width:100%"></div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">יום</div>
                        <div style="float:right;width:20%"><input type="number" min=1 max=5 name="theDay" id="theDay" style="width:100%"></div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">סדר</div>
                        <div style="float:right;width:20%"><input type="number" min=1 max=5 name="theOrder" id="theOrder" style="width:100%"></div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;margin-left:5%" onclick="doSubmit('splitForm')">שמירה</div>
                        <div style="float:right;width:20%;margin-left:5%">ביטול</div>
                    </div>
                    </form>
                </div>
                <%
                r.close
            elseif request("item")="session" then
                sql = "select * from [imknasim].[dbo].[session] where conference="&request("num")&" order by theOrder"
                r.open sql,strconn,1,3
                %>
                <div style="float:right;width:45%">
                <table style="font-size:1.5rem;padding:10px;margin:10px">
                <%
                if not r.eof then
                    %>
                    <tr style="border-bottom:1px solid #DADADA;height:50px"><td>שם</td><td>פיצול</td><td>האם להציג באג'נדה</td><td>סדר</td></tr>
                    <%
                    while not r.eof
                        thesplit="אין"
                        if r("split")<>0 then
                            sql1 = "select * from [imknasim].[dbo].[split] where theindex="&r("split")
                            r1.open sql1,strconn,1,3
                            if not r1.eof then thesplit=r1("theName")
                            r1.close
                        end if
                        %>
                        <tr style="border-bottom:1px solid #DADADA;height:50px"><td><%=r("theName")%></td><td><%=thesplit%></td><td><%=r("showinagenda")%></td><td><%=r("theOrder")%></td>
                            <td style="cursor:pointer" onclick="editItem('<%=request("item")%>',<%=r("theIndex")%>)">ערוך</td>
                            <td style="cursor:pointer" onclick="deleteItem('<%=request("item")%>',<%=r("theIndex")%>)">מחק</td>
                        </tr>
                        <%
                        r.movenext 
                    wend
                end if
                %>
                <tr><td style="cursor:pointer" colspan=6 onclick="editItem('<%=request("item")%>',0)">הוסף</td></tr>
                </table>
                </div>
                <div id="sessionEdit" style="float:right;width:50%;margin:10px;padding:10px;display:none;font-size:1.5rem">
                    <form id="sessionForm" method="post" action="saveTable.asp">
                    <input type="hidden" id="theConf" name="theConf" value="<%=request("num")%>">
                    <input type="hidden" id="theTable" name="theTable" value="session">
                    <input type="hidden" id="theNumber" name="theNumber">
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">שם</div>
                        <div style="float:right;width:50%"><input type="text" name="theName" id="theName" style="width:100%"></div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">פיצול</div>
                        <div style="float:right;width:50%">
                            <select name="split" id="split" style="width:100%">
                                <option value="0">בחר פיצול</option>
                                <%
                                sql1 = "select * from [imknasim].[dbo].[split] where theconference="&request("num")&" order by theOrder"
                                r1.open sql1,strconn,1,3
                                while not r1.eof
                                    %>
                                    <option value="<%=r1("theIndex")%>"><%=r1("theName")%></option>
                                    <%
                                    r1.movenext
                                wend
                                r1.close
                                %>
                            </select>
                        </div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">הצג באנג'נדה</div>
                        <div style="float:right;width:20%">
                            <select name="showinagenda" id="showinagenda" style="width:100%">
                                <option value="False">לא</option>
                                <option value="True">כן</option>
                            </select>
                        </div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">סדר</div>
                        <div style="float:right;width:20%"><input type="number" min=1 max=5 name="theOrder" id="theOrder" style="width:100%"></div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">יום</div>
                        <div style="float:right;width:20%"><input type="number" min=1 max=5 name="theDay" id="theDay" style="width:100%"></div>
                    </div>

                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">שם היו"ר</div>
                        <div style="float:right;width:50%"><input type="text" name="chairName" id="chairName" style="width:100%"></div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">תואר</div>
                        <div style="float:right;width:50%">
                            <select name="charTitle" id="chairTitle" style="width:100%">
                                <option value="0">בחר תואר</option>
                                <%
                                sql1 = "select * from [imknasim].[dbo].[title] order by theName"
                                r1.open sql1,strconn,1,3
                                while not r1.eof
                                    %>
                                    <option value="<%=r1("theIndex")%>"><%=r1("theName")%></option>
                                    <%
                                    r1.movenext
                                wend
                                r1.close
                                %>
                            </select>
                        </div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">תפקיד</div>
                        <div style="float:right;width:50%"><input type="text" name="chairJob" id="chairJob" style="width:100%"></div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">ארגון</div>
                        <div style="float:right;width:50%">
                            <select name="chairOrganization" id="charOrganization" style="width:100%">
                                <option value="0">בחר ארגון</option>
                                <%
                                sql1 = "select * from [imknasim].[dbo].[organization] order by theName"
                                r1.open sql1,strconn,1,3
                                while not r1.eof
                                    %>
                                    <option value="<%=r1("theIndex")%>"><%=r1("theName")%></option>
                                    <%
                                    r1.movenext
                                wend
                                r1.close
                                %>
                            </select>
                        </div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;margin-left:5%" onclick="doSubmit('sessionForm')">שמירה</div>
                        <div style="float:right;width:20%;margin-left:5%">ביטול</div>
                    </div>
                    </form>
                </div>
                <%
                r.close
            elseif request("item")="consponser" then
                sql = "select * from [imknasim].[dbo].[consponser] where conference="&request("num")&" order by theOrder"
                r.open sql,strconn,1,3
                %>
                <div style="float:right;width:45%">
                <table style="font-size:1.5rem;padding:10px;margin:10px">
                <%
                if not r.eof then
                    %>
                    <tr style="border-bottom:1px solid #DADADA;height:50px"><td>שם</td><td>סדר</td></tr>
                    <%
                    while not r.eof
                        thesponsor="לא ידוע"
                        if r("split")<>0 then
                            sql1 = "select * from [imknasim].[dbo].[sponsor] where theindex="&r("sponsor")
                            r1.open sql1,strconn,1,3
                            if not r1.eof then thesponsor=r1("theName")
                            r1.close
                        end if
                        %>
                        <tr style="border-bottom:1px solid #DADADA;height:50px"><td><%=thesponsor%></td><td><%=r("theOrder")%></td>
                            <td style="cursor:pointer" onclick="editItem('<%=request("item")%>',<%=r("theIndex")%>)">ערוך</td>
                            <td style="cursor:pointer" onclick="deleteItem('<%=request("item")%>',<%=r("theIndex")%>)">מחק</td>
                        </tr>
                        <%
                        r.movenext 
                    wend
                end if
                %>
                <tr><td style="cursor:pointer" colspan=6 onclick="editItem('<%=request("item")%>',0)">הוסף</td></tr>
                </table>
                </div>
                <div id="consponserEdit" style="float:right;width:50%;margin:10px;padding:10px;display:none;font-size:1.5rem">
                    <form id="consponserForm" method="post" action="saveTable.asp">
                    <input type="hidden" id="theConf" name="theConf" value="<%=request("num")%>">
                    <input type="hidden" id="theTable" name="theTable" value="consponser">
                    <input type="hidden" id="theNumber" name="theNumber">
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">בחר נותן חסות</div>
                        <div style="float:right;width:50%">
                            <select name="sponsor" id="sponsor" style="width:100%">
                                <option value="0">בחר נותן חסות</option>
                                <%
                                sql1 = "select * from [imknasim].[dbo].[sponsor] order by theName"
                                r1.open sql1,strconn,1,3
                                while not r1.eof
                                    %>
                                    <option value="<%=r1("theIndex")%>"><%=r1("theName")%></option>
                                    <%
                                    r1.movenext
                                wend
                                r1.close
                                %>
                            </select>
                        </div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">סדר</div>
                        <div style="float:right;width:20%"><input type="number" min=1 max=5 name="theOrder" id="theOrder" style="width:100%"></div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;margin-left:5%" onclick="doSubmit('consponserForm')">שמירה</div>
                        <div style="float:right;width:20%;margin-left:5%">ביטול</div>
                    </div>
                    </form>
                </div>
                <%
                r.close
            elseif request("item")="item" then
                sql = "select a.* from [imknasim].[dbo].[item] as a, [imknasim].[dbo].[session] as b where a.conference="&request("num")&" and a.session=b.theindex order by b.theOrder"
                r.open sql,strconn,1,3
                %>
                <div style="float:right;width:45%">
                <table style="font-size:1.5rem;padding:10px;margin:10px">
                <%
                if not r.eof then
                    %>
                    <tr style="border-bottom:1px solid #DADADA;height:50px"><td>שם</td><td>מושב</td><td>סוג</td><td>חסות</Td><td>משעה</td><td>עד שעה</td></tr>
                    <%
                    while not r.eof
                        theSession="אין"
                        if r("session")<>0 then
                            sql1 = "select * from [imknasim].[dbo].[session] where theindex="&r("session")
                            r1.open sql1,strconn,1,3
                            if not r1.eof then theSession=r1("theName")
                            r1.close
                        end if
                        sponsership="אין"
                        if r("sponsership")<>0 then
                            sql1 = "select * from [imknasim].[dbo].[sponsor] where theindex="&r("sponsership")
                            r1.open sql1,strconn,1,3
                            if not r1.eof then sponsership=r1("theName")
                            r1.close
                        end if
                        itemType="לא ידוע"
                        if r("itemType")<>0 then
                            sql1 = "select * from [imknasim].[dbo].[itemType] where theindex="&r("itemType")
                            r1.open sql1,strconn,1,3
                            if not r1.eof then itemType=r1("theName")
                            r1.close
                        end if
                        %>
                        <tr style="border-bottom:1px solid #DADADA;height:50px"><td><%=r("theName")%></td><td><%=theSession%></td><td><%=itemType%></td><td><%=sponsership%></td><td><%=r("fromHour")%></td><td><%=r("toHour")%></td>
                            <td style="cursor:pointer" onclick="editItem('<%=request("item")%>',<%=r("theIndex")%>)">ערוך</td>
                            <td style="cursor:pointer" onclick="deleteItem('<%=request("item")%>',<%=r("theIndex")%>)">מחק</td>
                        </tr>
                        <%
                        r.movenext 
                    wend
                end if
                %>
                <tr><td style="cursor:pointer" colspan=6 onclick="editItem('<%=request("item")%>',0)">הוסף</td></tr>
                </table>
                </div>
                <div id="itemEdit" style="float:right;width:50%;margin:10px;padding:10px;display:none;font-size:1.5rem">
                    <form id="itemForm" method="post" action="saveTable.asp">
                    <input type="hidden" id="theConf" name="theConf" value="<%=request("num")%>">
                    <input type="hidden" id="theTable" name="theTable" value="item">
                    <input type="hidden" id="theNumber" name="theNumber">
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">שם</div>
                        <div style="float:right;width:50%"><input type="text" name="theName" id="theName" style="width:100%"></div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">מושב</div>
                        <div style="float:right;width:50%">
                            <select name="session" id="session" style="width:100%">
                                <option value="0">בחר מושב</option>
                                <%
                                sql1 = "select * from [imknasim].[dbo].[session] where conference="&request("num")&" order by theOrder"
                                r1.open sql1,strconn,1,3
                                while not r1.eof
                                    %>
                                    <option value="<%=r1("theIndex")%>"><%=r1("theName")%></option>
                                    <%
                                    r1.movenext
                                wend
                                r1.close
                                %>
                            </select>
                        </div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">חסות</div>
                        <div style="float:right;width:50%">
                            <select name="sponsor" id="sponsor" style="width:100%">
                                <option value="0">בחר חסות</option>
                                <%
                                sql1 = "select * from [imknasim].[dbo].[sponsor] order by theName"
                                r1.open sql1,strconn,1,3
                                while not r1.eof
                                    %>
                                    <option value="<%=r1("theIndex")%>"><%=r1("theName")%></option>
                                    <%
                                    r1.movenext
                                wend
                                r1.close
                                %>
                            </select>
                        </div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">סוג</div>
                        <div style="float:right;width:50%">
                            <select name="itemType" id="itemType" style="width:100%">
                                <option value="0">בחר סוג</option>
                                <%
                                sql1 = "select * from [imknasim].[dbo].[itemType] order by theName"
                                r1.open sql1,strconn,1,3
                                while not r1.eof
                                    %>
                                    <option value="<%=r1("theIndex")%>"><%=r1("theName")%></option>
                                    <%
                                    r1.movenext
                                wend
                                r1.close
                                %>
                            </select>
                        </div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">סדר</div>
                        <div style="float:right;width:20%"><input type="number" min=1 max=5 name="theOrder" id="theOrder" style="width:100%"></div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">משעה</div>
                        <div style="float:right;width:20%"><input type="time" name="fromHour" id="fromHour" style="width:100%"></div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">עד שעה</div>
                        <div style="float:right;width:20%"><input type="time" name="toHour" id="toHour" style="width:100%"></div>
                    </div>
                    
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">שם המרצה</div>
                        <div style="float:right;width:50%"><input type="text" name="lecName" id="lecName" style="width:100%"></div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">תואר</div>
                        <div style="float:right;width:50%">
                            <select name="lecTitle" id="lecTitle" style="width:100%">
                                <option value="0">בחר תואר</option>
                                <%
                                sql1 = "select * from [imknasim].[dbo].[title] order by theName"
                                r1.open sql1,strconn,1,3
                                while not r1.eof
                                    %>
                                    <option value="<%=r1("theIndex")%>"><%=r1("theName")%></option>
                                    <%
                                    r1.movenext
                                wend
                                r1.close
                                %>
                            </select>
                        </div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">תפקיד</div>
                        <div style="float:right;width:50%"><input type="text" name="lecJob" id="lecJob" style="width:100%"></div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;width:10%">ארגון</div>
                        <div style="float:right;width:50%">
                            <select name="organization" id="organization" style="width:100%">
                                <option value="0">בחר ארגון</option>
                                <%
                                sql1 = "select * from [imknasim].[dbo].[organization] order by theName"
                                r1.open sql1,strconn,1,3
                                while not r1.eof
                                    %>
                                    <option value="<%=r1("theIndex")%>"><%=r1("theName")%></option>
                                    <%
                                    r1.movenext
                                wend
                                r1.close
                                %>
                            </select>
                        </div>
                    </div>
                    <div style="float:right;width:100%;margin-top:10px">
                        <div style="float:right;margin-left:5%" onclick="doSubmit('itemForm')">שמירה</div>
                        <div style="float:right;width:20%;margin-left:5%">ביטול</div>
                    </div>
                    </form>
                </div>
                <%
                r.close
            end if
            %>
        </div>
    </div>
</div>

<script>
function GoToPage(thePage,x,y)
{
	location=thePage+"?page="+x+"&cat="+y+"";
}
</script>