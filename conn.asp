<%@ LANGUAGE=VBScript CodePage=936%>
<%Response.Buffer=True%>
<%
ver="6.0"
dim conn
dim connstr
datapath    ="data/"
datafile    ="data.mdb"
Connstr="DBQ="&server.mappath(""&datapath&""&datafile&"")&";DRIVER={Microsoft Access Driver (*.mdb)};"
SqlNowString="Now()"
SqlChar="'"
Set conn=Server.CreateObject("ADODB.Connection")
conn.open ConnStr
%>