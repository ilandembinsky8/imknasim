<!-- #include file="secure.inc" -->
<script>
    function refreshPage(x){
        <% 
        if request("page")="[stamps].[dbo].[MainTable]" then
        %>
        location.href="inner.asp?page=[stamps].[dbo].[MainTable]&year="+x+"";
        <%
        else
        %>
        location.href="inner.asp?page=[stamps].[dbo].[vCancels]&year="+x+"";
        <%
        end if
        %>
    }
    function refreshPage1(x){
        location.href="inner.asp?page=[stamps].[dbo].[list]&current="+x+"";
    }
</script>
<%
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")


thetable = trim(request("page"))
currpage = thetable
select case thetable
    case "[imknasim].[dbo].[conference]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""
    case "[imknasim].[dbo].[itemtype]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""
    case "[imknasim].[dbo].[organization]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""
    case "[imknasim].[dbo].[orgtype]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""
    case "[imknasim].[dbo].[sponsor]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""
    case "[imknasim].[dbo].[title]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""
    case "[imknasim].[dbo].[section]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""
    case "[imknasim].[dbo].[fieldtype]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""    
    case "[imknasim].[dbo].[fields]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""    
    case "[imknasim].[dbo].[choicetable]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""    
    case "[imknasim].[dbo].[sex]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""    
    case "[imknasim].[dbo].[country]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""   
    case "[imknasim].[dbo].[validationtype]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""    
    case "[imknasim].[dbo].[profession]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""   
    case "[imknasim].[dbo].[participanttype]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""    
    case "[imknasim].[dbo].[participation]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""    
    case "[imknasim].[dbo].[registertype]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""   
    case "[imknasim].[dbo].[paymenttype]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""    
    case "[imknasim].[dbo].[paymentstatus]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""
    case "[imknasim].[dbo].[roomtype]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""    
    case "[imknasim].[dbo].[nutrition]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""    
    case "[imknasim].[dbo].[accesstype]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""    
    case "[imknasim].[dbo].[documents]"
	    var_num = "theindex"
	    var_title = "thename"
	    var_title1 = ""    
end select



if thetable<>"" then
    if thetable="[imknasim].[dbo].[fields]" then
        sql = "SELECT "&var_num&", "&var_title&" FROM "&thetable
        if request("item")&"a"<>"a" then sql = sql & " where fieldsection="&request("item")
        sql =  sql & " order by "&var_title
    elseif var_title1="" then
		sql = "SELECT "&var_num&", "&var_title&" FROM "&thetable&" order by "&var_title
    else
        sql = "SELECT "&var_num&", "&var_title&", "&var_title1&" FROM "&thetable&" order by "&var_title
	end if
    r.open sql,strconn,1,3
    allFields = r.Fields.Count
end if

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
    <%
    if thetable<>"" then
    if thetable="[imknasim].[dbo].[fields]" then
        if request("item")&"a"="a" then
            theItem=0
        else
            theItem=request("item")
        end if
        %>
        <div style="float:right;width:100%">
        <div style="float:right;width:12.5%"><input name="section" value="0" type="radio" checked onclick="location.href='inner.asp?page=<%=request("page")%>'">הכל</div>
        <%
        sql1 = "select * from section order by theOrder"
        r1.open sql1,strconn,1,3
        while not r1.eof
           %>
           <div style="float:right;width:12.5%"><input name="section" value="<%=r1("theIndex")%>" type="radio" onclick="location.href='inner.asp?page=<%=request("page")%>&item=<%=r1("theIndex")%>'" <%if r1("theIndex")=int(theItem) then response.write "checked"%>><%=r1("theName")%></div>
           <% 
           r1.movenext
        wend
        r1.close
        %>
        </div>
        <%
    end if

    while not r.eof
        
        response.Write "<a href='edit.asp?page="&thetable&"&num="&r(var_num)&"&all="&allFields&"' style='font-size:24px;line-height:135%;direction:rtl'>"&r(var_title)
        if var_title1<>"" then response.write "<span style='color:white'>--</span>"&r(var_title1)
        if var_title2<>"" then response.write "<span style='color:white'>--</span>"&r(var_title2)
        response.write "</a>&nbsp;&nbsp;&nbsp;"
        if request("page")="[imknasim].[dbo].[conference]" then
            response.write "<a style='font-size:24px;line-height:135%;direction:rtl;font-weight:bold;' href='forumDetails.asp?page="&thetable&"&num="&r(var_num)&"'>פרטים</a>&nbsp;&nbsp;&nbsp;"
            response.write "<a style='font-size:24px;line-height:135%;direction:rtl;font-weight:bold;' href='../agenda.html?num="&r(var_num)&"' target='_blank'>צפייה</a>&nbsp;&nbsp;&nbsp;"
        end if
        response.write "<a style='font-size:24px;line-height:135%;direction:rtl;font-weight:bold;' href='duplicate.asp?page="&thetable&"&num="&r(var_num)&"&all="&allFields&"'>שכפל</a>&nbsp;&nbsp;&nbsp;"
        
        response.write "<br>"
        
        r.MoveNext
    wend
    
    end if%>
    <input type=button style="width:200px;height:40px;background-color:#EAEAEA;box-sizing:border-box;border-radius:8px;border:0px;margin-top:20px;font-size:24px" value="הוסף" onclick="location.href='add.asp?page=<%=thetable%>'">
    </div>
</div>

<script>
function GoToPage(thePage,x,y)
{
	location=thePage+"?page="+x+"&cat="+y+"";
}
</script>