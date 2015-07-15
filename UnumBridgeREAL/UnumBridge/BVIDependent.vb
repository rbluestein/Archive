Imports UnumBridge.biz.plane.www

Public Class BVIDependent
    Private Common As New Common
    Private DBase As New DBase

    Public Function SetDependentValues(ByRef UnumSession As UnumSession) As DependentsDependent()
        Dim i As Integer
        Dim dt As DataTable

        dt = DBase.GetConvertDTSql("SELECT * FROM EmpDependent WHERE EmpID = '" & UnumSession.EmpID & "' AND Active=1", UnumSession.ServerIPAddress, UnumSession.DBName)

        If dt.Rows.Count = 0 Then
            Dim depdep() As DependentsDependent
            Return depdep
        Else
            Dim depdep(dt.Rows.Count - 1) As DependentsDependent
            For i = 0 To dt.Rows.Count - 1
                depdep(i) = GetDependent(dt, i)
            Next
            Return depdep
        End If
    End Function

    Private Function GetDependent(ByVal dt, ByVal i) As DependentsDependent
        Dim dep As New DependentsDependent

        dep.Address1 = Common.DoTrim(dt.Rows(i)("AddrLine1").ToText, 50)
        dep.Address2 = Common.DoTrim(dt.Rows(i)("AddrLine2").ToText, 50)
        dep.City = Common.DoTrim(dt.Rows(i)("City").ToText, 50)

        Common.SetDateValues(dt.Rows(i)("BirthDate").ToText, dep.DateOfBirth, dep.DateOfBirthSpecified)

        Common.SetDependentStatus(dt.Rows(i)("HandicappedInd").ToText, dt.Rows(i)("StudentInd").ToText, dep.DependentStatus, dep.DependentStatusSpecified)
        dep.EMail = ""
        dep.FirstName = Common.DoTrim(dt.Rows(i)("FirstName").ToText, 50)
        Common.SetGenderValues(dt.Rows(i)("SexCode").ToText, dep.Gender, dep.GenderSpecified)
        dep.HomePhone = ""
        dep.LastName = Common.DoTrim(dt.Rows(i)("LastName").ToText, 50)
        dep.MaritalStatus = DependentsDependentMaritalStatus.Unknown
        dep.MaritalStatusSpecified = False
        dep.MiddleName = Common.DoTrim(dt.Rows(i)("MI").ToText, 25)
        dep.Occupation = ""
        Common.SetRelationTypeValues(dt.Rows(i)("RelationCode").ToText, dt.Rows(i)("SexCode").ToText, dep.RelationType, dep.RelationTypeSpecified)
        dep.SchoolName = Common.DoTrim(dt.Rows(i)("SchoolName").ToText, 100)
        dep.SSN = dt.Rows(i)("SSN").ToText
        Common.SetStateValues(dt.Rows(i)("State").ToText, dep.State, dep.StateSpecified)
        dep.Title = ""
        dep.WorkPhone = ""
        dep.Zip = Common.GetFullZipCode(dt.Rows(i)("Zip5").ToText, dt.Rows(i)("Zip4").ToText)
        Return dep
    End Function

    Private Function SetTestDependents() As DependentsDependent()
        Dim i As Integer
        Dim dt As New DataTable
        Dim depdep(dt.Rows.Count) As DependentsDependent

        ' For i = 0 To dt.Rows.Count - 1
        depdep(0) = GetTestDependent(dt, i)
        ' Next
        Return depdep
    End Function

    Private Function GetTestDependent(ByVal dt, ByVal i) As DependentsDependent
        Dim dep As New DependentsDependent

        dep.Address1 = ""
        dep.Address2 = ""
        dep.City = ""
        dep.DateOfBirth = DateTime.Parse("1/1/1970")
        dep.DateOfBirthSpecified = True
        dep.DependentStatus = DependentStatusList.NA
        dep.DependentStatusSpecified = False
        dep.EMail = ""
        dep.FirstName = "Test"
        dep.Gender = DependentsDependentGender.Male
        dep.GenderSpecified = True
        dep.HomePhone = ""
        dep.LastName = "Tester"
        dep.MaritalStatus = DependentsDependentMaritalStatus.Single
        dep.MaritalStatusSpecified = True
        dep.MiddleName = ""
        dep.Occupation = ""
        dep.RelationType = DependentRelationType.Spouse
        dep.RelationTypeSpecified = True
        dep.SchoolName = ""
        dep.SSN = ""
        dep.State = DependentStateList.AL
        dep.StateSpecified = True
        dep.Title = ""
        dep.WorkPhone = ""
        dep.Zip = ""
        Return dep
    End Function
End Class
