Public Class LaunchUnumSession
    Inherits System.Web.UI.Page

    Private Common As New Common
    Private DBase As New DBase
    Protected WithEvents phDisplayRegID As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents PlaceHolder1 As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents phErrorMsg As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents phDisplayErrorMsg As System.Web.UI.WebControls.PlaceHolder
    Private UnumSession As UnumSession

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try

            UnumSession = Session("UnumSession")
            phDisplayErrorMsg.Visible = False
            phDisplayRegID.Visible = False

            If 0 = 0 Then
                UnumSession.DBName = Request.QueryString("DBName")
                UnumSession.ServerIPAddress = Request.QueryString("ServerIPAddress")
                UnumSession.EmpID = Request.QueryString("EmpID")
                UnumSession.EnrollerID = Request.QueryString("EnrollerID")
                UnumSession.ActivityID = Request.QueryString("ActivityID")

            Else
                UnumSession.DBName = "HT"
                UnumSession.ServerIPAddress = "192.168.1.15"
                'UnumSession.EmpID = "100596"  'freeman
                UnumSession.EmpID = "101058"  ' michael adair
                'UnumSession.EmpID = "100211"  'coco
                'UnumSession.EnrollerID = "cwright"
                UnumSession.EnrollerID = "aearly"
                UnumSession.ActivityID = "63fb66f9-d6bf-41b5-a6d3-b818fc37d536"
            End If

            UnumSession.GetClientVariables()

            If Not Page.IsPostBack Then
                GetUnumUserID()
                If UnumSession.EnrollerIsUnumRegistered Then
                    Response.Redirect("LaunchUnumSessionConnect.aspx")
                Else
                    phDisplayRegID.Visible = True
                End If
            Else
                ProcessRegistration()
                Response.Redirect("LaunchUnumSessionConnect.aspx")
            End If

        Catch ex As Exception
            phDisplayErrorMsg.Visible = True
        End Try
    End Sub

    Private Sub GetUnumUserID()
        Dim dt As DataTable
        Dim Sql As String
        Dim QueryPack As CmdAsst.QueryPack
        Dim UnumUserID As String

        Sql = "SELECT UnumUserID FROM IAMS..UnumUserIDLookup WHERE BVIUserID = '" & UnumSession.EnrollerID & "'"
        QueryPack = DBase.GetDTWithQueryPack(Sql, UnumSession.ServerIPAddress, "IAMS")
        If QueryPack.Success Then
            If QueryPack.dt.rows.Count = 0 Then
                UnumSession.EnrollerIsUnumRegistered = False
            Else
                UnumSession.EnrollerIsUnumRegistered = True
                UnumSession.UnumEnrollerID = QueryPack.dt.Rows(0)(0)
            End If
        Else
            Throw New Exception("Unable to read from UnumUserIDLookup table")
        End If
    End Sub

    Private Sub ProcessRegistration()
        Dim dt As DataTable
        Dim Sql As String
        Dim QueryPack As CmdAsst.QueryPack

        Sql = "INSERT INTO IAMS..UnumUserIDLookup (BVIUserID, UnumUserID) Values ('" & UnumSession.EnrollerID & "', '" & Request.Form("txtUnumUserID") & "')"
        QueryPack = DBase.GetDTWithQueryPack(Sql, UnumSession.ServerIPAddress, "IAMS")
        If QueryPack.Success Then
            UnumSession.EnrollerIsUnumRegistered = True
            UnumSession.UnumEnrollerID = Request.Form("txtUnumUserID")
        Else
            Throw New Exception("Unable to write to UnumUserIDLookup table")
        End If
    End Sub
End Class
