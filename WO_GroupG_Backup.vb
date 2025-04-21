
Imports Microsoft.MetadirectoryServices
Imports System.Text
Imports System.Security.Cryptography

Public Class MAExtensionObject
    Implements IMASynchronization


    Public Sub New()

    End Sub

    Public Sub Initialize() Implements IMASynchronization.Initialize
        ' TODO: Add initialization code here
    End Sub

    Public Sub Terminate() Implements IMASynchronization.Terminate
        ' TODO: Add termination code here
    End Sub

    Public Function ShouldProjectToMV(ByVal csentry As CSEntry, ByRef MVObjectType As String) As Boolean Implements IMASynchronization.ShouldProjectToMV

        If csentry("OBJEKT_TYPE").StringValue.Contains("Medlem-E") Or csentry("OBJEKT_TYPE").StringValue.Contains("Medlem-N") Then
            MVObjectType = "person"
            ShouldProjectToMV = True
        End If
        'Throw New EntryPointNotImplementedException()

    End Function

    Public Function FilterForDisconnection(ByVal csentry As CSEntry) As Boolean Implements IMASynchronization.FilterForDisconnection
        ' TODO: Add connector filter code here
        Throw New EntryPointNotImplementedException()
    End Function

    Public Sub MapAttributesForJoin(ByVal FlowRuleName As String, ByVal csentry As CSEntry, ByRef values As ValueCollection) Implements IMASynchronization.MapAttributesForJoin
        Select Case FlowRuleName
            Case "puUserJoin"

                If csentry("AKTORID").IsPresent Then

                    'If csentry("AKTORID").IsPresent Then
                    Dim AccountName As String = csentry("AKTORID").Value
                    values.Add("PU-M" & AccountName)
                End If
                'Lagt til 29.08.2019 for å hindre duplikatproblematikk
            Case "xxEmployeeID"
                If csentry("AKTORID").IsPresent Then
                    Dim xxEmployeeID As String = "PUM" & csentry("AKTORID").Value
                    values.Add(xxEmployeeID)
                End If
        End Select
        'Throw New EntryPointNotImplementedException()
    End Sub

    Public Function ResolveJoinSearch(ByVal joinCriteriaName As String, ByVal csentry As CSEntry, ByVal rgmventry() As MVEntry, ByRef imventry As Integer, ByRef MVObjectType As String) As Boolean Implements IMASynchronization.ResolveJoinSearch
        Throw New EntryPointNotImplementedException()
    End Function

    Public Sub MapAttributesForImport(ByVal FlowRuleName As String, ByVal csentry As CSEntry, ByVal mventry As MVEntry) Implements IMASynchronization.MapAttributesForImport
        Select Case FlowRuleName

            Case "WO_GroupName"
                If csentry("FORENINGSNAVN").IsPresent And csentry("OBJEKT_TYPE").IsPresent And csentry("OBJEKT_TYPE").Value.Contains("Lokallagsleder") Then 'And csentry("OBJEKT_TYPE").Value <> "Ansatt" Then
                    Dim Foreningsnavn As String = csentry("FORENINGSNAVN").Value
                    Foreningsnavn = Foreningsnavn.Replace("Utdanningsforbundet", "")
                    Foreningsnavn = Foreningsnavn.Replace("ø", "o")
                    Foreningsnavn = Foreningsnavn.Replace("æ", "a")
                    Foreningsnavn = Foreningsnavn.Replace("å", "a")
                    Foreningsnavn = Foreningsnavn.Replace("Ø", "O")
                    Foreningsnavn = Foreningsnavn.Replace("Æ", "A")
                    Foreningsnavn = Foreningsnavn.Replace("Å", "A")
                    Foreningsnavn = Foreningsnavn.Replace("\s+", "")
                    Foreningsnavn = Foreningsnavn.Replace(" ", "")
                    Foreningsnavn = Foreningsnavn.Replace("(", "_")
                    Foreningsnavn = Foreningsnavn.Replace(")", "")
                    mventry("WO_GroupName").Value = "PU.App.EPI_WEL" + Foreningsnavn

                End If

            Case "puConnectedGroups"
                'Adding person to PU.Org.Members 
                'If csentry("OBJEKT_TYPE").IsPresent And csentry("OBJEKT_TYPE").StringValue.Contains("Medlem") And csentry("OBJEKT_TYPE").Value <> "Ansatt" Then
                'Dim Gruppenavn As String = "PU.Org.Members"
                'mventry("puConnectedGroups").Values.Add(Gruppenavn)
                'End If

                'Adding person to tillitsvalgtbedriftgruppen if Tillitsvalgtbedrift = E,N
                If csentry("OBJEKT_TYPE").IsPresent And csentry("OBJEKT_TYPE").StringValue.Contains("TillitsvalgtBedrift-E") Or csentry("OBJEKT_TYPE").StringValue.Contains("TillitsvalgtBedrift-N") Then 'And csentry("OBJEKT_TYPE").Value <> "Ansatt"
                    Dim Gruppenavn As String = "PU.App.EPI_WinorgTillitsvalgtBedrift"
                    mventry("puConnectedGroups").Values.Add(Gruppenavn)
                End If

                'If tillitsvalgtbedrift = S and already member of "PU.App.EPI_WinorgTillitsvalgtBedrift" - remove
                If csentry("OBJEKT_TYPE").StringValue.Contains("TillitsvalgtBedrift-S") Then
                    Try
                        Dim Gruppenavn As String = "PU.App.EPI_WinorgTillitsvalgtBedrift"
                        mventry("puConnectedGroups").Values.Remove(Gruppenavn)
                    Catch ex As AttributeDoesNotExistOnObjectException
                    End Try
                End If

                'Adding person to folkeskolelaggruppen if folkeskolelag = E,N
                If csentry("OBJEKT_TYPE").IsPresent And csentry("OBJEKT_TYPE").StringValue.Contains("Folkeskolelag-E") Or csentry("OBJEKT_TYPE").StringValue.Contains("Folkeskolelag-N") Then 'And csentry("OBJEKT_TYPE").Value <> "Ansatt" Then
                    Dim Gruppenavn As String = "PU.App.EPI_WinorgFolkeskolelag"
                    mventry("puConnectedGroups").Values.Add(Gruppenavn)
                End If

                'If folkeskolelaggruppen = S and already member of "PU.App.EPI_WinorgFolkeskolelag" - remove
                If csentry("OBJEKT_TYPE").StringValue.Contains("Folkeskolelag-S") Then
                    Try
                        Dim Gruppenavn As String = "PU.App.EPI_WinorgFolkeskolelag"
                        mventry("puConnectedGroups").Values.Remove(Gruppenavn)
                    Catch ex As AttributeDoesNotExistOnObjectException
                    End Try
                End If

                'Adding person to tillitsvalgtgruppen if Tillitsvalgt = E,N
                If csentry("OBJEKT_TYPE").IsPresent And csentry("OBJEKT_TYPE").StringValue.Contains("Tillitsvalgt-E") Or csentry("OBJEKT_TYPE").StringValue.Contains("Tillitsvalgt-N") Then 'And csentry("OBJEKT_TYPE").Value <> "Ansatt" Then
                    Dim Gruppenavn As String = "PU.App.EPI_WinorgTillitsvalgt"
                    mventry("puConnectedGroups").Values.Add(Gruppenavn)
                End If
                'if tillitsvalgt = S and already a member of PU.App.EPI_WinorgTillitsvalg - remove
                If csentry("OBJEKT_TYPE").Value.Contains("Tillitsvalgt-S") Then
                    Try
                        Dim Gruppenavn As String = "PU.App.EPI_WinorgTillitsvalgt"
                        mventry("puConnectedGroups").Values.Remove(Gruppenavn)
                    Catch ex As AttributeDoesNotExistOnObjectException
                    End Try
                End If

                'Adding person to lokallagsleder groups if Lokallagsleder = E,N 
                If csentry("FORENINGSNAVN").IsPresent And csentry("OBJEKT_TYPE").IsPresent And csentry("OBJEKT_TYPE").StringValue.Contains("Lokallagsleder-E") Or csentry("OBJEKT_TYPE").StringValue.Contains("Lokallagsleder-N") Then 'And csentry("OBJEKT_TYPE").Value <> "Ansatt" Then
                    Dim Foreningsnavn As String = csentry("FORENINGSNAVN").Value
                    Foreningsnavn = Foreningsnavn.Replace("Utdanningsforbundet", "")
                    Foreningsnavn = Foreningsnavn.Replace("ø", "o")
                    Foreningsnavn = Foreningsnavn.Replace("æ", "a")
                    Foreningsnavn = Foreningsnavn.Replace("å", "a")
                    Foreningsnavn = Foreningsnavn.Replace("Ø", "O")
                    Foreningsnavn = Foreningsnavn.Replace("Æ", "A")
                    Foreningsnavn = Foreningsnavn.Replace("Å", "A")
                    Foreningsnavn = Foreningsnavn.Replace("\s+", "")
                    Foreningsnavn = Foreningsnavn.Replace(" ", "")
                    Foreningsnavn = Foreningsnavn.Replace("(", "_")
                    Foreningsnavn = Foreningsnavn.Replace(")", "")
                    mventry("puConnectedGroups").Values.Add("PU.App.EPI_WEL" + Foreningsnavn)
                    mventry("puConnectedGroups").Values.Add("PU.App.EPI_WinorgLokallagsleder")
                    mventry("puConnectedGroups").Values.Add("PU.App.EPI_ServerWebEditors")
                    mventry("puConnectedGroups").Values.Add("PU.App.EPI_ServerImageVaultEditors")
                    mventry("puConnectedGroups").Values.Add("PU.App.EPI_WELokallag")
                End If

                'Removing person from lokallagsleder groups if Lokallagsleder = S and puConnectedgroups contains foreningsnavn
                If csentry("FORENINGSNAVN").IsPresent And csentry("OBJEKT_TYPE").IsPresent And csentry("OBJEKT_TYPE").StringValue.Contains("Lokallagsleder-S") Then
                    Try
                        Dim Foreningsnavn As String = csentry("FORENINGSNAVN").Value
                        Foreningsnavn = Foreningsnavn.Replace("Utdanningsforbundet", "")
                        Foreningsnavn = Foreningsnavn.Replace("ø", "o")
                        Foreningsnavn = Foreningsnavn.Replace("æ", "a")
                        Foreningsnavn = Foreningsnavn.Replace("å", "a")
                        Foreningsnavn = Foreningsnavn.Replace("Ø", "O")
                        Foreningsnavn = Foreningsnavn.Replace("Æ", "A")
                        Foreningsnavn = Foreningsnavn.Replace("Å", "A")
                        Foreningsnavn = Foreningsnavn.Replace("\s+", "")
                        Foreningsnavn = Foreningsnavn.Replace(" ", "")
                        Foreningsnavn = Foreningsnavn.Replace("(", "_")
                        Foreningsnavn = Foreningsnavn.Replace(")", "")
                        Foreningsnavn = "PU.App.EPI_WEL" + Foreningsnavn
                        Dim Groups As ValueCollection
                        Groups = mventry("puConnectedGroups").Values
                        Groups.Remove(Foreningsnavn)
                        Groups.Remove("PU.App.EPI_WinorgLokallagsleder")
                        Groups.Remove("PU.App.EPI_ServerWebEditors")
                        Groups.Remove("PU.App.EPI_ServerImageVaultEditors")
                        Groups.Remove("PU.App.EPI_WELokallag")
                        mventry("puConnectedGroups").Values = Groups

                        'mventry("puConnectedGroups").Values.Remove(Foreningsnavn)
                        'mventry("puConnectedGroups").Values.Remove("PU.App.EPI_WinorgLokallagsleder")
                        'mventry("puConnectedGroups").Values.Remove("PU.App.EPI_ServerWebEditors")
                        'mventry("puConnectedGroups").Values.Remove("PU.App.EPI_ServerImageVaultEditors")
                        'mventry("puConnectedGroups").Values.Remove("PU.App.EPI_WELokallag")

                    Catch ex As AttributeDoesNotExistOnObjectException
                    End Try
                End If

                If csentry("AKTORID").IsPresent And csentry("MEDLEM_STATUS").IsPresent Then
                    'If csentry("MEDLEM_STATUS").Value = "N" Or csentry("MEDLEM_STATUS").Value = "E" Then
                    Try
                            Dim Gruppenavn As String = "PU.Org.Members"
                            mventry("puConnectedGroups").Values.Add(Gruppenavn)
                        Catch ex As AttributeDoesNotExistOnObjectException
                        End Try
                    End If
                'End If

            Case "WinOrg"
                If csentry("OBJEKT_STATUS").IsPresent Then
                    mventry("Winorg").BooleanValue = True
                Else
                    mventry("Winorg").BooleanValue = False
                End If

            Case "WO_DisplayName" 'setting displayname
                If csentry("OBJEKT_TYPE").StringValue.Contains("Lokallagsleder") And csentry("OBJEKT_TYPE").Value IsNot "Ansatt" Then
                    Dim Foreningsnavn As String = csentry("FORENINGSNAVN").Value
                    Foreningsnavn = Foreningsnavn.Replace("Utdanningsforbundet", "")
                    Foreningsnavn = Foreningsnavn.Replace("ø", "o")
                    Foreningsnavn = Foreningsnavn.Replace("æ", "a")
                    Foreningsnavn = Foreningsnavn.Replace("å", "a")
                    Foreningsnavn = Foreningsnavn.Replace("Ø", "O")
                    Foreningsnavn = Foreningsnavn.Replace("Æ", "A")
                    Foreningsnavn = Foreningsnavn.Replace("Å", "A")
                    Foreningsnavn = Foreningsnavn.Replace("\s+", "")
                    Foreningsnavn = Foreningsnavn.Replace(" ", "")
                    Foreningsnavn = Foreningsnavn.Replace("(", "_")
                    Foreningsnavn = Foreningsnavn.Replace(")", "")
                    mventry("displayName").Value = "WO_PU.App.EPI_WEL" + Foreningsnavn

                Else
                    'If csentry("Foreningsnavn").IsPresent And csentry("MEDLEM_STATUS").Value Is "S" Then
                    'mventry("description").Value = ""
                End If

            Case "WO_UserAccountName"
                If csentry("AKTORID").IsPresent Then
                    Dim AktorID As String = csentry("AktorID").Value
                    mventry("accountName").Value = "PU-M" + AktorID
                    'mventry("puConnectedGroups").Values.Add("PU.Org.Members")
                Else
                    ' If csentry("Foreningsnavn").IsPresent And csentry("MEDLEM_STATUS").Value Is "S" Then
                    '  mventry("accountname").Value = ""
                    'End If
                End If

            Case "WO_Tillitsvalgt"
                If csentry("OBJEKT_TYPE").StringValue.Contains("Tillitsvalgt-E") Or csentry("OBJEKT_TYPE").StringValue.Contains("Tillitsvalgt-N") Then
                    mventry("WO_Tillitsvalgt").BooleanValue = True
                Else
                    mventry("WO_Tillitsvalgt").BooleanValue = False
                End If

            Case "WO_TillitsvalgtBedrift"
                If csentry("OBJEKT_TYPE").StringValue.Contains("TillitsvalgtBedrift-E") Or csentry("OBJEKT_TYPE").StringValue.Contains("TillitsvalgtBedrift-N") Then
                    mventry("WO_TillitsvalgtBedrft").BooleanValue = True
                Else
                    mventry("WO_TillitsvalgtBedrft").BooleanValue = False
                End If

            Case "WO_Lokallagsleder"
                If csentry("OBJEKT_TYPE").StringValue.Contains("Lokallagsleder-E") Or csentry("OBJEKT_TYPE").StringValue.Contains("Lokallagsleder-N") Then
                    mventry("WO_Lokallagsleder").BooleanValue = True
                Else
                    mventry("WO_Lokallagsleder").BooleanValue = False
                End If

            Case "WO_Folkeskolelag"
                If csentry("OBJEKT_TYPE").StringValue.Contains("Folkeskolelag-E") Or csentry("OBJEKT_TYPE").StringValue.Contains("Folkeskolelag-N") Then
                    mventry("WO_Folkeskolelag").BooleanValue = True
                Else
                    mventry("WO_Folkeskolelag").BooleanValue = False
                End If

            Case "PUxxEmployeeID"
                If csentry("AKTORID").IsPresent Then
                    mventry("xxEmployeeID").Value = "PUM" & csentry("AKTORID").Value
                End If
            Case "HashPersonalID"
                If csentry("PERSNR").IsPresent Then
                    Dim InputString As String = csentry("PERSNR").Value
                    Dim sha256 As SHA256 = SHA256Managed.Create()
                    Dim bytes As Byte() = Encoding.UTF8.GetBytes(InputString)
                    Dim hash As Byte() = sha256.ComputeHash(bytes)
                    Dim stringBuilder As New StringBuilder()

                    For i As Integer = 0 To hash.Length - 1
                        stringBuilder.Append(hash(i).ToString("X2"))
                    Next

                    mventry("i2-NationalID").Value = stringBuilder.ToString()
                End If

            Case "Mobile"
                If csentry("MBT").Value.StartsWith("00") Then
                    Dim Mobile1 As String = csentry("MBT").Value
                    Mobile1 = Mobile1.Remove(0, 2)
                    mventry("mobile").Value = "+" & Mobile1
                ElseIf csentry("MBT").Value.StartsWith("+47") Then
                    mventry("mobile").Value = csentry("MBT").Value
                ElseIf csentry("MBT").Value.StartsWith("+") Then
                    mventry("mobile").Value = csentry("MBT").Value
                Else
                    mventry("mobile").Value = "+47" & csentry("MBT").Value
                End If

            Case "PUMedlemStatus"
                If csentry("MEDLEM_STATUS").IsPresent Then
                    mventry("PUMedlemStatus").Value = csentry("MEDLEM_STATUS").Value
                Else
                    mventry("PUMedlemStatus").Value = "S"
                End If

            Case "puBirthdate"
                If csentry("PERSNR").IsPresent Then

                    Dim Persnr As String = csentry("PERSNR").Value
                    Dim Extract As String = Left(Persnr, 6)
                    If Persnr.Length >= 11 Then
                        Dim Year As String = Right(Extract, 2)
                        Dim Month As String = Mid(Extract, 3, 2)
                        Dim Day As String = Left(Extract, 2)
                        Dim IndividSifre As String = Mid(Persnr, 7, 3)
                        Dim newString As Int32 = Convert.ToInt32(IndividSifre)
                        Dim IntYear As Int32 = Convert.ToInt32(Year)
                        'Test for D number
                        Dim YearCorrection As Int32 = Convert.ToInt32(Year)
                        Dim HMonth As Int32 = Convert.ToInt32(Month)
                        Dim HMonthCorrection As Int32 = 40
                        Dim DDay As Int32 = Convert.ToInt32(Day)
                        Dim DDayCorrection As Int32 = 40
                        If DDay >= 40 Then
                            DDay = (DDay) - (DDayCorrection)
                            If DDay < 10 Then
                                Dim PadZero As String
                                PadZero = "0" & DDay
                                Day = PadZero
                            Else
                                Dim NoPad As String
                                NoPad = DDay.ToString
                                Day = NoPad
                            End If

                        End If
                        'Else
                        '    Day = Day
                        'Day = DDay.ToString()

                        If HMonth >= 40 Then
                            HMonth = (HMonth) - (HMonthCorrection)
                            If HMonth < 10 Then
                                Dim PadZero As String
                                PadZero = "0" & HMonth
                                Month = PadZero
                                '    End If
                                'Else Month = Month
                            End If
                        End If
                        If YearCorrection < 10 Then
                            Dim PadYear As String
                            PadYear = "200" & YearCorrection.ToString()
                            Year = PadYear
                            mventry("i2-birthdate").Value = Year & Mid(Extract, 3, 2) & Day & "000000.0Z"
                        End If
                        If YearCorrection > 10 Then
                            Dim PadYear As String
                            PadYear = "19" & YearCorrection.ToString()
                            Year = PadYear
                            mventry("i2-birthdate").Value = Year & Mid(Extract, 3, 2) & Day & "000000.0Z"
                        End If

                    End If

                        'Select Case newString
                        '        Case 0 To 499
                        '            Dim FormatedYear As String
                        '            FormatedYear = "19" & Year
                        '            mventry("i2-birthdate").Value = FormatedYear & Mid(Extract, 3, 2) & Day & "000000.0Z"
                        '        Case 900 To 999
                        '            Dim FormatedYear As String
                        '            FormatedYear = "19" & Year
                        '            mventry("i2-birthdate").Value = FormatedYear & Mid(Extract, 3, 2) & Day & "000000.0Z"
                        '        Case 500 To 749
                        '            Select Case IntYear
                        '                Case 54 To 99
                        '                    Dim FormatedYear As String
                        '                    FormatedYear = "18" & Year
                        '                    mventry("i2-birthdate").Value = FormatedYear & Mid(Extract, 3, 2) & Day & "000000.0Z"
                        '            End Select
                        '        Case 500 To 999
                        '            Select Case IntYear
                        '                Case 0 To 40
                        '                    Dim FormatedYear As String
                        '                    FormatedYear = "20" & Year
                        '                    mventry("i2-birthdate").Value = FormatedYear & Mid(Extract, 3, 2) & Day & "000000.0Z"
                        '            End Select
                        '    End Select
                    End If


            Case "UserPrinicipalName"
                If csentry("AKTORID").IsPresent Then
                    mventry("userPrincipalName").Value = "PU-M" & csentry("AKTORID").Value & "@udf.no"
                End If

            Case "displayName"
                If csentry("FORNAVN").IsPresent And csentry("ETTERNAVN").IsPresent Then
                    mventry("displayName").Value = csentry("FORNAVN").Value.Trim & " " & csentry("ETTERNAVN").Value.Trim
                End If
                'Added for primaryGroupID for all pu-m users to be set to 644960 (PU.Org.Members)
            Case "primaryGroupID"
                If csentry("AKTORID").IsPresent Then
                    Dim primaryGroupID As Int32 = 644960
                    mventry("primaryGroupID").IntegerValue = primaryGroupID
                End If
            Case "firstName"
                If csentry("FORNAVN").IsPresent Then
                    mventry("firstName").Value = csentry("FORNAVN").Value.Trim
                End If
            Case "lastName"
                If csentry("ETTERNAVN").IsPresent Then
                    mventry("lastName").Value = csentry("ETTERNAVN").Value.Trim
                End If
        End Select

        ' Throw New EntryPointNotImplementedException()
    End Sub

    Public Sub MapAttributesForExport(ByVal FlowRuleName As String, ByVal mventry As MVEntry, ByVal csentry As CSEntry) Implements IMASynchronization.MapAttributesForExport
        Throw New EntryPointNotImplementedException()
    End Sub

    Public Function Deprovision(ByVal csentry As CSEntry) As DeprovisionAction Implements IMASynchronization.Deprovision
        ' TODO: Remove this throw statement if you implement this method
        Throw New EntryPointNotImplementedException()
    End Function


End Class
