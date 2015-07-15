<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Return.aspx.vb" Inherits="UnumBridge._Return" ValidateRequest="False"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Unum Return</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK title="BVIStyle" href="BVI.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body>
		<form id="Form1" action="Return.aspx" method="post" runat="server">
			<table class="ReturnTbl" style="LEFT: 140px; POSITION: absolute; TOP: 14px" cellSpacing="0"
				cellPadding="0" width="500" border="0">
				<tr>
					<asp:literal id="litPrompt" runat="server" EnableViewState="False"></asp:literal></tr>
				<tr>
					<td class="CellSeparator"></td>
				</tr>
			</table>
		</form>
		<script language="javascript">
			function fnClose()  {				
				window.opener= "x"
				window.close()
			}
		</script>
	</body>
</HTML>
