<%@ LANGUAGE="VBSCRIPT" %>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1">
<TITLE>Document Title</TITLE>
</HEAD>
<BODY>

<!-- Insert HTML here -->
<%
dim hwnd
dim donn
hwnd=Request.Form("hwnd")
donn=Request.Form("donn")
%>
Ducer Connection completed
donn:<% =donn %>
hwnd:<% =hwnd %>
</BODY>
</HTML>
