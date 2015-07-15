<%@ Page Language="vb" AutoEventWireup="false" Codebehind="LaunchUnumSessionConnect.aspx.vb" Inherits="UnumBridge.LaunchUnumSessionConnect" EnableViewState="False" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
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
			function fnSubmit()  {
				if (form1.hdGoToUnum.value == '1')  {
					var frm = document.getElementById("form1")
					frm.setAttribute("action", form1.hdExchangeURL.value)
					form1.submit()				
				}		
			}
		</script>
	</body>
</HTML>
