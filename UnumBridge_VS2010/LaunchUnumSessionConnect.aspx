<%@ Page Language="VB" AutoEventWireup="false" CodeFile="LaunchUnumSessionConnect.aspx.vb" Inherits="LaunchUnumSessionConnect" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
	<HEAD>
		<title>Launch Unum Session</title>
	</HEAD>
	<body onload="fnSubmit()">
		<form id="form1" method="post" action="LaunchUnumSessionConnect.aspx">
			<asp:Literal id="litExchangeSessionID" runat="server" EnableViewState="False"></asp:Literal>
			<asp:Literal id="litExchangeURL" runat="server" EnableViewState="False"></asp:Literal>
			<asp:Literal id="litGoToUnum" runat="server" EnableViewState="False"></asp:Literal>			
			<asp:Label id="lblMsg" EnableViewState="False" runat="server"></asp:Label></form>
		<script language="javascript">
		    function fnSubmit() {
		        if (form1.hdGoToUnum.value == '1') {
		            var frm = document.getElementById("form1")
		            frm.setAttribute("action", form1.hdExchangeURL.value)
		            form1.submit()
		        }
		    }
		</script>
	</body>
</HTML>
