Imports System.Data

Public Enum BenefitTypeEnum
    VWBAccident = 5
End Enum

Public Enum ESBenefitTypeEnum
    VWBAccident = 6
End Enum

Public Enum PolicyScopeEnum
    Individual = 1
    Multiple = 2
End Enum

Public Class UnumSession
    Private cSessionID As String
    Private cEmpID As String
    Private cEmpSSN As String
    Private cPayFrequencyDesc As String
    Private cEnrollerID As String
    Private cServerIPAddress As String
    Private cDBName As String
    Private cActivityID As String
    Private cBVIProductID As String
    Private cClientID As String
    Private cESBenefitType As ESBenefitTypeEnum
    Private cBenefitType As BenefitTypeEnum
    Private cMessage As String = String.Empty
    Private cLastLoginIP As String
    Private cPayFrequencyCode As String
    Private cPayFrequencyFactor As Integer
    Private cUnumEnrollerID As String
    Private cEnrollerIsUnumRegistered As Boolean

    Public Property LastLoginIP()
        Get
            Return cLastLoginIP
        End Get
        Set(ByVal Value)
            cLastLoginIP = Value
        End Set
    End Property
    Public Property Message()
        Get
            Return cMessage
        End Get
        Set(ByVal Value)
            cMessage = Value
        End Set
    End Property

    Public ReadOnly Property ESBenefitType()
        Get
            Return cESBenefitType
            'Select Case cDBName
            '    Case "HT"
            '        Return ESBenefitTypeEnum.VWBAccident
            'End Select
        End Get
    End Property

    Public ReadOnly Property BenefitType()
        Get
            Return cBenefitType
            'Select Case cDBName
            '    Case "HT"
            '        Return BenefitTypeEnum.VWBAccident
            'End Select
        End Get
    End Property
    Public ReadOnly Property BVIProductID()
        Get
            Return cBVIProductID
            'Select Case cDBName
            '    Case "HT"
            '        Return "VWB-Accident"
            'End Select
        End Get
    End Property
    Public ReadOnly Property ClientID()
        Get
            Return cClientID
            'Select Case cDBName
            '    Case "HT"
            '        Return "HarrisTeeter"
            'End Select
        End Get
    End Property
    Public Property PayFrequencyDesc()
        Get
            Return cPayFrequencyDesc
        End Get
        Set(ByVal Value)
            cPayFrequencyDesc = Value
        End Set
    End Property

    Public Property EnrollerID()
        Get
            Return cEnrollerID
        End Get
        Set(ByVal Value)
            cEnrollerID = Value
        End Set
    End Property

    Public Property DBName()
        Get
            Return cDBName
        End Get
        Set(ByVal Value)
            cDBName = Value
        End Set
    End Property

    Public Property ActivityID()
        Get
            Return cActivityID
        End Get
        Set(ByVal Value)
            cActivityID = Value
        End Set
    End Property

    Public Property ServerIPAddress()
        Get
            Return cServerIPAddress
        End Get
        Set(ByVal Value)
            cServerIPAddress = Value
        End Set
    End Property

    Public Property SessionID()
        Get
            Return cSessionID
        End Get
        Set(ByVal Value)
            cSessionID = Value
        End Set
    End Property

    Public Property EmpID()
        Get
            Return cEmpID
        End Get
        Set(ByVal Value)
            cEmpID = Value
        End Set
    End Property

    Public Property EmpSSN()
        Get
            Return cEmpSSN
        End Get
        Set(ByVal Value)
            cEmpSSN = Value
        End Set
    End Property

    Public ReadOnly Property ecURL()
        Get
            'Return "https://acceptance.plane.biz/exchange2/exchangeconnect.asmx"    ' test
            'Return "https://demo.plane.biz/exchange2/exchangeconnect.asmx"              ' demo
            Return "https://www.plane.biz/exchange2/exchangeconnect.asmx"                ' prod
        End Get
    End Property

    Public ReadOnly Property ecGUID() As String
        Get
            'Return "47ba2e96-4613-454e-800c-bb42ef674a31"   ' test
            Return "b9689063-986f-4517-af1c-d7cb7890fe70"    '  prod
        End Get
    End Property

    Public Property PayFrequencyCode()
        Get
            Return cPayFrequencyCode
        End Get
        Set(ByVal Value)
            cPayFrequencyCode = Value
        End Set
    End Property

    Public Property PayFrequencyFactor()
        Get
            Return cPayFrequencyFactor
        End Get
        Set(ByVal Value)
            cPayFrequencyFactor = Value
        End Set
    End Property

    Public Property UnumEnrollerID()
        Get
            Return cUnumEnrollerID
        End Get
        Set(ByVal Value)
            cUnumEnrollerID = Value
        End Set
    End Property

    Public Property EnrollerIsUnumRegistered()
        Get
            Return cEnrollerIsUnumRegistered
        End Get
        Set(ByVal Value)
            cEnrollerIsUnumRegistered = Value
        End Set
    End Property

    Public Sub GetClientVariables()
        Dim DBase As New DBase
        Dim CmdAsst As New CmdAsst(DBase, CommandType.StoredProcedure, "usp_UnumBridge", cServerIPAddress, "IAMS")
        Dim QueryPack As CmdAsst.QueryPack
        CmdAsst.AddVarChar("DBName", cDBName, 50)
        QueryPack = CmdAsst.Execute()
        cBVIProductID = QueryPack.dt.rows(0)("BVIProductID")
        cClientID = QueryPack.dt.rows(0)("ClientID")

        Select Case cDBName
            Case "HT"
                Select Case QueryPack.dt.rows(0)("BenefitType")
                    Case "VWBAccident"
                        cBenefitType = BenefitTypeEnum.VWBAccident
                End Select
                Select Case QueryPack.dt.rows(0)("ESBenefitType")
                    Case "VWBAccident"
                        cESBenefitType = ESBenefitTypeEnum.VWBAccident
                End Select
        End Select

    End Sub


End Class
