<!-- #include file="function.asp" -->
<%
dim count
id=int(Request("id"))
set rs=server.createobject("ADODB.Recordset")
sql="select * from [list] where id="&id
rs.Open sql,Conn,1,3
rs("count")=rs("count")+1
rs.update
count=rs("count")
rs.close
Set rs=Nothing
%>
document.write('<%=count%>');