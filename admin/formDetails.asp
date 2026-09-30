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
        <div style="float:right;width:100%;height:40px;font-size:2rem">
            <%
            sql = "select * from [imknasim].[dbo].[section] order by theOrder"
            r.open sql,strconn,1,3
            while not r.eof
                found=false
                sql1 = "select * from [imknasim].[dbo].[conffield] as a, [imknasim].[dbo].[fields] as b where a.thefield=b.theindex and b.fieldsection="&r("theindex")&" and a.theconf="&request("num")
                r1.open sql1,strconn,1,3
                if not r1.eof then found=true
                r1.close
                %>
                <div style="float:right;width:20%"><input id="section<%=r("theIndex")%>" type="checkbox" <%if found=true then response.write "checked"%>>&nbsp;<%=r("theName")%></div>
                <%
                r.movenext
            wend
            r.close
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