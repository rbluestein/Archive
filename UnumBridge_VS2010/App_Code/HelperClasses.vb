Imports UnumBridge.biz.plane.www

Public Class AppPolSummary
    Public FirstName As String
    Public MiddleName As String
    Public LastName As String
    Public CoverageBeginDate As DateTime
    Public ProductID As String
    Public MonthlyPremium As Decimal
    Public PayPeriodPremium As Decimal
    Public CoverageAmount As Decimal
    Public CoverageAmountSpecified As Boolean
End Class

Public Class AppPolData
    Public PersonSSN As String
    Public Description As String
    Public MonthlyPremium As Decimal
    Public PayPeriodPremium As Decimal
    Public CoverageAmount As Decimal
    Public CoverageAmountSpecified As Boolean
End Class

Public Class SaveResults
    Public AppID As String
    Public Success As Boolean
    Public TechErrMsg As String
    Public Message As String
End Class

Public Class BenefitObj
    Private cColl As New Collection


    ' ----------------------------------------------------------------------
    ' Sample 1:
    ' VWB-Accident    Individual
    '
    ' Sample 2:
    ' VWB-Accident    Multiple
    ' UL                     Individual
    ' -----------------------------------------------------------------------

    Public ReadOnly Property Coll()
        Get
            Return cColl
        End Get
    End Property

    Public Sub Add(ByVal ProductID As String, ByVal PolicyScope As PolicyScopeEnum)
        cColl.Add(New Item(ProductID, PolicyScope))
    End Sub

    Public Function ProductIDExists(ByVal ProductID As String) As Boolean
        Dim i As Integer
        For i = 1 To cColl.Count
            'If cColl(i).ProductID = ProductID AndAlso cColl(i).PolicyScope = PolicyScopeEnum.Multiple Then
            If cColl(i).ProductID = ProductID Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Class Item
        Private cProductID As String
        Private cPolicyScope As PolicyScopeEnum

        Public ReadOnly Property ProductID()
            Get
                Return cProductID
            End Get
        End Property

        Public ReadOnly Property PolicyScope()
            Get
                Return cPolicyScope
            End Get
        End Property

        Public Sub New(ByVal ProductID As String, ByVal PolicyScope As PolicyScopeEnum)
            cProductID = ProductID
            cPolicyScope = PolicyScope
        End Sub
    End Class
End Class