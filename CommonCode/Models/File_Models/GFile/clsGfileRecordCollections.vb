'clsGFileRecords
'Purpose:       To hold all record collections from a G File
'Updates:
'Note: 
'               Public members & Functions PascalCase
'               private members & local variables camelCase
'Created on:    2023/08/04 
'By:            Grant Haynes

Option Strict On
Option Explicit On

Imports System.Collections.Generic

Public Class GFileRecordCollections

    'RecordsA
    'Purpose:       To hold the collection of A records and their associated methods
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/12/22
    'By:            Grant Haynes

    Public Class RecordsA

        Public Property Records As New List(Of GFileRecords.RecordA)

        'Default Constructor
        Public Sub New()

        End Sub

    End Class

    'RecordsB
    'Purpose:       To hold the collection of B records and their associated methods
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/12/22
    'By:            Grant Haynes

    Public Class RecordsB

        Public Property Records As New List(Of GFileRecords.RecordB)

        'Default Constructor
        Public Sub New()

        End Sub

        Public Function GetRecordByDateInRange(SearchDate As Date) As GFileRecords.RecordB
            For Each recordB As GFileRecords.RecordB In Records
                If recordB.ProjectStartDate.Value IsNot Nothing And recordB.ProjectEndDate.Value IsNot Nothing Then
                    If ((SearchDate >= CDate(recordB.ProjectStartDate.Value)) And (SearchDate <= CDate(recordB.ProjectEndDate.Value))) Then
                        Return recordB
                    End If
                End If
            Next
            Return New GFileRecords.RecordB()
        End Function

        Public Function GetRecordByStartDate(SearchDate As Date) As GFileRecords.RecordB
            For Each recordB As GFileRecords.RecordB In Records
                If recordB.ProjectStartDate.Value IsNot Nothing Then
                    If SearchDate = CDate(recordB.ProjectStartDate.Value) Then
                        Return recordB
                    End If
                End If
            Next
            Return New GFileRecords.RecordB()
        End Function

        Public Function GetRecordByEndDate(SearchDate As Date) As GFileRecords.RecordB
            For Each recordB As GFileRecords.RecordB In Records
                If recordB.ProjectEndDate.Value IsNot Nothing Then
                    If SearchDate <= CDate(recordB.ProjectEndDate.Value) Then
                        Return recordB
                    End If
                End If
            Next
            Return New GFileRecords.RecordB()
        End Function

        Public Function GetRecordByStartOrEndDate(SearchDate As Date) As GFileRecords.RecordB
            For Each recordB As GFileRecords.RecordB In Records
                If recordB.ProjectStartDate.Value IsNot Nothing And recordB.ProjectEndDate.Value IsNot Nothing Then
                    If SearchDate = CDate(recordB.ProjectStartDate.Value) Or SearchDate = CDate(recordB.ProjectEndDate.Value) Then
                        Return recordB
                    End If
                End If
            Next
            Return New GFileRecords.RecordB()
        End Function

    End Class

    'clsRecordsC
    'Purpose:       To hold the collection of C records and their associated methods
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/12/22
    'By:            Grant Haynes

    Public Class RecordsC

        Public Property Records As New List(Of GFileRecords.RecordC)

        'Default Constructor
        Public Sub New()

        End Sub

        Public Function GetRecordsbyOriginSSNDifferentialSSNandSessionID(OriginSSN As Integer, DifferentialSSN As Integer, SessionID As Integer) As GFileRecords.RecordC
            For Each record As GFileRecords.RecordC In Records
                If record.OriginSSN.Value = OriginSSN And record.DifferentialSSN.Value = DifferentialSSN And record.SessionID = SessionID Then
                    Return record
                End If
            Next
            Return New GFileRecords.RecordC()
        End Function

        Public Function GetRecordsBySessionID(SessionID As Integer) As List(Of GFileRecords.RecordC)
            Dim returnRecords As New List(Of GFileRecords.RecordC)
            For Each recordC As GFileRecords.RecordC In Records
                If recordC.SessionID = SessionID Then
                    returnRecords.Add(recordC)
                End If
            Next
            Return returnRecords
        End Function

    End Class

    'RecordsD
    'Purpose:       To hold the collection of D records and their associated methods
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/12/22
    'By:            Grant Haynes

    Public Class RecordsD

        Public Property Records As New List(Of GFileRecords.RecordD)

        'Default Constructor
        Public Sub New()

        End Sub

        Public Function GetRecordsBySessionID(SessionID As Integer) As List(Of GFileRecords.RecordD)
            Dim returnRecords As New List(Of GFileRecords.RecordD)
            For Each recordD As GFileRecords.RecordD In Records
                If recordD.SessionID = SessionID Then
                    returnRecords.Add(recordD)
                End If
            Next
            Return returnRecords
        End Function

    End Class

    'RecordsE
    'Purpose:       To hold the collection of E records and their associated methods
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/12/22
    'By:            Grant Haynes

    Public Class RecordsE

        Public Property Records As New List(Of GFileRecords.RecordE)

        'Default Constructor
        Public Sub New()

        End Sub

        Public Function GetRecordsBySessionID(SessionID As Integer) As List(Of GFileRecords.RecordE)
            Dim returnRecords As New List(Of GFileRecords.RecordE)
            For Each recordE As GFileRecords.RecordE In Records
                If recordE.SessionID = SessionID Then
                    returnRecords.Add(recordE)
                End If
            Next
            Return returnRecords
        End Function

    End Class

    'RecordsF
    'Purpose:       To hold the collection of F records and their associated methods
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/12/22
    'By:            Grant Haynes

    Public Class RecordsF

        Public Property Records As New List(Of GFileRecords.RecordF)

        'Default Constructor
        Public Sub New()

        End Sub

        Public Function GetRecordsBySessionID(SessionID As Integer) As List(Of GFileRecords.RecordF)
            Dim returnRecords As New List(Of GFileRecords.RecordF)
            For Each recordF As GFileRecords.RecordF In Records
                If recordF.SessionID = SessionID Then
                    returnRecords.Add(recordF)
                End If
            Next
            Return returnRecords
        End Function

    End Class

    'RecordsG
    'Purpose:       To hold the collection of G records and their associated methods
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/12/22
    'By:            Grant Haynes

    Public Class RecordsG

        Public Property Records As New List(Of GFileRecords.RecordG)

        'Default Constructor
        Public Sub New()

        End Sub

        Public Function GetRecordsBySessionID(SessionID As Integer) As List(Of GFileRecords.RecordG)
            Dim returnRecords As New List(Of GFileRecords.RecordG)
            For Each recordG As GFileRecords.RecordG In Records
                If recordG.SessionID = SessionID Then
                    returnRecords.Add(recordG)
                End If
            Next
            Return returnRecords
        End Function

    End Class

    'RecordsH
    'Purpose:       To hold the collection of H records and their associated methods
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/12/22
    'By:            Grant Haynes

    Public Class RecordsH

        Public Property Records As New List(Of GFileRecords.RecordH)

        'Default Constructor
        Public Sub New()

        End Sub

        Public Function GetRecordsBySessionID(SessionID As Integer) As List(Of GFileRecords.RecordH)
            Dim returnRecords As New List(Of GFileRecords.RecordH)
            For Each recordH As GFileRecords.RecordH In Records
                If recordH.SessionID = SessionID Then
                    returnRecords.Add(recordH)
                End If
            Next
            Return returnRecords
        End Function

        Public Function GetRecordBySSN(SSN As Integer) As GFileRecords.RecordH
            For Each recordH As GFileRecords.RecordH In Records
                If recordH.StationSerialNumber.Value = SSN Then
                    Return recordH
                End If
            Next
            Return New GFileRecords.RecordH
        End Function

    End Class

    'RecordsI
    'Purpose:       To hold the collection of I records and their associated methods
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/12/22
    'By:            Grant Haynes

    Public Class RecordsI

        Public Property Records As New List(Of GFileRecords.RecordI)

        'Default Constructor
        Public Sub New()

        End Sub

        Public Function GetRecordsBySessionID(SessionID As Integer) As List(Of GFileRecords.RecordI)
            Dim returnRecords As New List(Of GFileRecords.RecordI)
            For Each recordI As GFileRecords.RecordI In Records
                If recordI.SessionID = SessionID Then
                    returnRecords.Add(recordI)
                End If
            Next
            Return returnRecords
        End Function

    End Class

End Class