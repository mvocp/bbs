<%
myfilename="cookie.asp"
dim adminusername, adminuserpass
if Request.Cookies("adminusername")=empty then outadmin
if Request.Cookies("adminuserpass")=empty then outadmin
If conn.Execute("Select id From [user] where distinction=10 and username='"&Request.Cookies("adminusername")&"' and userpass='"&Request.Cookies("adminuserpass")&"' " ).eof Then outadmin
sub outadmin
Response.Cookies("adminusername")=""
Response.Cookies("adminuserpass")=""
set rs=nothing
set rs1=nothing
set conn=nothing
response.write "<meta http-equiv=refresh content=0;url=./admin.asp>"
Response.End
end sub
%>