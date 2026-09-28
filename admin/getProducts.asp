

    <!-- #include file = "secure.inc" -->
    {["items":[
    <%
    sql = "select * from [stamps].[dbo].[products]"
    r.open sql,strconn,1,3
    while not r.eof
        %>
        {"theName":"<%=r("theName")%>","description":"<%=r("description")%>","theImage":"<%=r("theImage")%>","thePrice":"<%=r("thePrice")%>"}
        <%
        r.movenext
        if not r.eof then response.write ","
    wend
    r.close
    response.write "]}"
    %>
