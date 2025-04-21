
Imports Microsoft.MetadirectoryServices

Public Class MAExtensionObject
    Implements IMASynchronization

    Public Sub Initialize() Implements IMASynchronization.Initialize
        ' TODO: Add initialization code here
    End Sub

    Public Sub Terminate() Implements IMASynchronization.Terminate
        ' TODO: Add termination code here
    End Sub

    Public Function ShouldProjectToMV(ByVal csentry As CSEntry, ByRef MVObjectType As String) As Boolean Implements IMASynchronization.ShouldProjectToMV

        If csentry("OBJEKT_TYPE").StringValue.Contains("Medlem") Then
            MVObjectType = "person"
            ShouldProjectToMV = True
        End If

        If csentry("OBJEKT_TYPE").StringValue.Contains("Tillitsvalgt") Then
            MVObjectType = "group"
            ShouldProjectToMV = True
        End If
        If csentry("OBJEKT_TYPE").StringValue.Contains("TillitsvalgtBedrift") Then

            MVObjectType = "group"
            ShouldProjectToMV = True
        End If
        If csentry("OBJEKT_TYPE").StringValue.Contains("Folkeskolelag") Then
            MVObjectType = "group"
            ShouldProjectToMV = True

        End If


        'Throw New EntryPointNotImplementedException()

    End Function

    Public Function FilterForDisconnection(ByVal csentry As CSEntry) As Boolean Implements IMASynchronization.FilterForDisconnection
        ' TODO: Add connector filter code here
        Throw New EntryPointNotImplementedException()
    End Function

    Public Sub MapAttributesForJoin(ByVal FlowRuleName As String, ByVal csentry As CSEntry, ByRef values As ValueCollection) Implements IMASynchronization.MapAttributesForJoin
        ' TODO: Add join mapping code here
        Select Case FlowRuleName
            Case "puUserJoin"
                Dim AccountName As String = csentry("AKTORID").Value
                If csentry("AKTORID").IsPresent Then
                    values.Add("PU-M" & AccountName)
                End If

            'Case "puGroupJoin"
                '    Select Case csentry("OBJEKT_TYPE").Value
                'Case "Lokallagsleder"
                '   Dim Foreningsnavn As String = csentry("FORENINGSNAVN").Value
                '  If csentry("OBJEKT_TYPE").StringValue.Contains("Lokallagsleder") And csentry("OBJEKT_TYPE").Value IsNot "Ansatt" Then
                ' Foreningsnavn = Foreningsnavn.Replace("Utdanningsforbundet", "")
                'Foreningsnavn = Foreningsnavn.Replace("ø", "o")
                'Foreningsnavn = Foreningsnavn.Replace("æ", "a")
                'Foreningsnavn = Foreningsnavn.Replace("å", "a")
                'Foreningsnavn = Foreningsnavn.Replace("\s+", "")
                'Foreningsnavn = Foreningsnavn.Replace(" ", "")
                'Foreningsnavn = Foreningsnavn.Replace("(", "_")
                'Foreningsnavn = Foreningsnavn.Replace(")", "")
                'values.Add("PU.App.EPI_WEL" & Foreningsnavn)
                'End If
                'Case "TillitsvalgtBedrift"
                '   Dim TillitsvalgtBedrift As String = "PU.App.EPI_WinorgTillitsvalgtBedrift"
                '  If csentry("OBJEKT_TYPE").StringValue.Contains("TillitsvalgtBedrift") And csentry("OBJEKT_TYPE").Value IsNot "Ansatt" Then
                ' values.Add(TillitsvalgtBedrift)
                'End If
                'Case "Lokallagsleder"
                '   Dim Lokallagsleder As String = "PU.App.EPI_WinorgLokallagsleder"
                '  If csentry("OBJEKT_TYPE").StringValue.Contains("Lokallagsleder") And csentry("OBJEKT_TYPE").Value IsNot "Ansatt" Then
                ' values.Add(Lokallagsleder)
                'End If
                'Case "Folkeskolelag"
                '   Dim Folkeskolelag As String = "PU.App.EPI_WinorgFolkeskolelag"
                '  If csentry("OBJEKT_TYPE").StringValue.Contains("Folkeskolelag") And csentry("OBJEKT_TYPE").Value IsNot "Ansatt" Then
                ' values.Add(Folkeskolelag)
                'End If
            Case "Tillitsvalgt"
                        Dim Tillitsvalgt As String = "PU.App.EPI_WinorgTillitsvalgt"
                        If csentry("OBJEKT_TYPE").StringValue.Contains("Tillitsvalgt") And csentry("OBJEKT_TYPE").Value IsNot "Ansatt" Then
                            values.Add(Tillitsvalgt)
                        End If
                End Select
        'End Select

        ''Throw New EntryPointNotImplementedException()
    End Sub

    Public Function ResolveJoinSearch(ByVal joinCriteriaName As String, ByVal csentry As CSEntry, ByVal rgmventry() As MVEntry, ByRef imventry As Integer, ByRef MVObjectType As String) As Boolean Implements IMASynchronization.ResolveJoinSearch
        'ResolveJoinSearch = False
        'imventry = -1

        'If csentry("OBJEKT_STATUS").StringValue.Contains("Tillitsvalgt") Then
        '    Dim Tillitsvalgt As String
        '    Dim mventry As MVEntry
        '    Tillitsvalgt = "PU.App.EPI_WinorgTillitsvalgt"

        '    For Each mventry In rgmventry
        '        If mventry("puGruppe").IsPresent AndAlso
        '            mventry("puGrupppe").Value.Equals(Tillitsvalgt) Then
        '            ResolveJoinSearch = True

        '            Exit Function
        '        End If
        Throw New EntryPointNotImplementedException()
        '    Next
        'End If
    End Function

    Public Sub MapAttributesForImport(ByVal FlowRuleName As String, ByVal csentry As CSEntry, ByVal mventry As MVEntry) Implements IMASynchronization.MapAttributesForImport
        Select Case FlowRuleName

            Case "WO_GroupName"
                If csentry("FORENINGSNAVN").IsPresent And csentry("OBJEKT_TYPE").IsPresent And csentry("OBJEKT_TYPE").Value.Contains("Lokallagsleder") And csentry("OBJEKT_TYPE").Value <> "Ansatt" Then
                    Dim Foreningsnavn As String = csentry("FORENINGSNAVN").Value
                    Foreningsnavn = Foreningsnavn.Replace("Utdanningsforbundet", "")
                    Foreningsnavn = Foreningsnavn.Replace("ø", "o")
                    Foreningsnavn = Foreningsnavn.Replace("æ", "a")
                    Foreningsnavn = Foreningsnavn.Replace("å", "a")
                    Foreningsnavn = Foreningsnavn.Replace("\s+", "")
                    Foreningsnavn = Foreningsnavn.Replace(" ", "")
                    Foreningsnavn = Foreningsnavn.Replace("(", "_")
                    Foreningsnavn = Foreningsnavn.Replace(")", "")
                    mventry("WO_GroupName").Value = "PU.App.EPI_WEL" + Foreningsnavn
                End If

            Case "puGruppeTillitsvalgtBedrift"
                If csentry("OBJEKT_TYPE").IsPresent And csentry("OBJEKT_TYPE").Value.Contains("TillitsvalgtBedrift") And csentry("OBJEKT_TYPE").Value <> "Ansatt" Then
                    Dim Gruppenavn As String = "PU.App.EPI_WinorgTillitsvalgtBedrift"
                    mventry("puGruppetillitsvalgtBedrift").Value = Gruppenavn
                End If

            'Case "puGruppetillitsvalgtBedriftMembers"
            '    If csentry("OBJEKT_TYPE").IsPresent And csentry("OBJEKT_TYPE").Value.Contains("TillitsvalgtBedrift") And csentry("AKTORID").IsPresent And csentry("OBJEKT_TYPE").Value <> "Ansatt" Then
            '        Dim GruppeMedlem As String = "PU-M" & csentry("AKTORID").Value
            '        mventry("puGruppetillitsvalgtBedriftMembers").Values.Add(GruppeMedlem)
            '    End If


            Case "puGruppeFolkeskolelag"
                If csentry("OBJEKT_TYPE").IsPresent And csentry("OBJEKT_TYPE").Value.Contains("Folkeskolelag") And csentry("OBJEKT_TYPE").Value <> "Ansatt" Then
                    Dim Gruppenavn As String = "PU.App.EPI_WinorgFolkeskolelag"
                    mventry("puGruppeFolkeskolelag").Value = Gruppenavn
                End If

            'Case "puGruppeFolkeskolelagMembers"
            '    If csentry("OBJEKT_TYPE").IsPresent And csentry("OBJEKT_TYPE").Value.Contains("Folkeskolelag") And csentry("AKTORID").IsPresent And csentry("OBJEKT_TYPE").Value <> "Ansatt" Then
            '        Dim GruppeMedlem As String = "PU-M" & csentry("AKTORID").Value
            '        mventry("puGruppeFolkeskolelagMembers").Values.Add(GruppeMedlem)
            '    End If

            Case "puGruppeTillitsvalgt"
                If csentry("OBJEKT_TYPE").IsPresent And csentry("OBJEKT_TYPE").Value.Contains("Tillitsvalgt") And csentry("OBJEKT_TYPE").Value <> "Ansatt" Then
                    Dim Gruppenavn As String = "PU.App.EPI_WinorgTillitsvalgt"
                    mventry("puGruppetillitsvalgt").Value = Gruppenavn
                End If

            Case "puGruppeTillitsvalgtMembers"
                If csentry("OBJEKT_TYPE").IsPresent And csentry("OBJEKT_TYPE").Value.Contains("Tillitsvalgt") And csentry("AKTORID").IsPresent And csentry("OBJEKT_TYPE").Value <> "Ansatt" Then
                    mventry("puGruppeTillitsvalgtMembers").Values.Add(csentry("AKTORID").Value)
                End If

            Case "WO_DisplayName" 'setting displayname
                If csentry("OBJEKT_TYPE").StringValue.Contains("Lokallagsleder") And csentry("OBJEKT_TYPE").Value IsNot "Ansatt" Then
                    Dim Foreningsnavn As String = csentry("FORENINGSNAVN").Value
                    Foreningsnavn = Foreningsnavn.Replace("Utdanningsforbundet", "")
                    Foreningsnavn = Foreningsnavn.Replace("ø", "o")
                    Foreningsnavn = Foreningsnavn.Replace("æ", "a")
                    Foreningsnavn = Foreningsnavn.Replace("å", "a")
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
                Else
                    ' If csentry("Foreningsnavn").IsPresent And csentry("MEDLEM_STATUS").Value Is "S" Then
                    '  mventry("accountname").Value = ""
                    'End If
                End If

            Case "WO_Tillitsvalgt"
                If csentry("OBJEKT_TYPE").StringValue.Equals("Folkeskolelag,Lokallagsleder,Medlem,Tillitsvalgt,TillitsvalgtBedrift") Or csentry("OBJEKT_TYPE").StringValue.Equals("Lokallagsleder,Medlem,Tillitsvalgt") Or csentry("OBJEKT_TYPE").StringValue.Equals("Lokallagsleder,Medlem,Tillitsvalgt,TillitsvalgtBedrift") Or csentry("OBJEKT_TYPE").StringValue.Equals("Medlem,Tillitsvalgt,TillitsvalgtBedrift") Or csentry("OBJEKT_TYPE").StringValue.Equals("Medlem,Tillitsvalgt") Then
                    mventry("WO_Tillitsvalgt").BooleanValue = True
                Else
                    mventry("WO_Tillitsvalgt").BooleanValue = False
                End If

            Case "WO_TillitsvalgtBedrift"
                If csentry("OBJEKT_TYPE").StringValue.Contains("TillitsvalgtBedrift") Then
                    mventry("WO_TillitsvalgtBedrft").BooleanValue = True
                Else
                    mventry("WO_TillitsvalgtBedrft").BooleanValue = False
                End If

            Case "WO_Lokallagsleder"
                If csentry("OBJEKT_TYPE").StringValue.Contains("Lokallagsleder") Then
                    mventry("WO_Lokallagsleder").BooleanValue = True
                Else
                    mventry("WO_Lokallagsleder").BooleanValue = False
                End If

            Case "WO_Folkeskolelag"
                If csentry("OBJEKT_TYPE").StringValue.Contains("Folkeskolelag") Then
                    mventry("WO_Folkeskolelag").BooleanValue = True
                Else
                    mventry("WO_Folkeskolelag").BooleanValue = False
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
