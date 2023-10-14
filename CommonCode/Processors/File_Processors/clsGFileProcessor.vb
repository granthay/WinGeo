'clsGFileProcessor
'Purpose:       To provide methods to process a gfile
'Updates:
'Note: 
'               Public members & Functions PascalCase
'               private members & local variables camelCase
'Created on:    2023/08
'By:            Grant Haynes

Option Strict On
Option Explicit On

Imports System.Collections.Generic

Public Class GFileProcessor

    Public Property Gfile As New GFile()

    'Default Constructor
    Public Sub New()

    End Sub

    Public Sub New(Gfile As GFile)
        _gfile = Gfile
    End Sub

    Public Function CreateVectorCorrelationMatrix(SessionID As Integer) As Double(,)
        Dim relatedRecordsC As List(Of GFileRecords.RecordC) = _Gfile.RecordsC.GetRecordsBySessionID(SessionID)
        Dim returnMatrix(relatedRecordsC.Count() * 3, relatedRecordsC.Count() * 3) As Double
        For Each recordD As GFileRecords.RecordD In _Gfile.RecordsD.GetRecordsBySessionID(SessionID)
            If recordD.RowIndexNumber1.Value IsNot Nothing And recordD.ColumnIndexNumber1.Value IsNot Nothing And recordD.Correlation1.Value IsNot Nothing Then
                returnMatrix(CInt(recordD.RowIndexNumber1.Value), CInt(recordD.ColumnIndexNumber1.Value)) = CDbl(recordD.Correlation1.Value)
            End If
            If recordD.RowIndexNumber2.Value IsNot Nothing And recordD.ColumnIndexNumber2.Value IsNot Nothing And recordD.Correlation2.Value IsNot Nothing Then
                returnMatrix(CInt(recordD.RowIndexNumber2.Value), CInt(recordD.ColumnIndexNumber2.Value)) = CDbl(recordD.Correlation2.Value)
            End If
            If recordD.RowIndexNumber3.Value IsNot Nothing And recordD.ColumnIndexNumber3.Value IsNot Nothing And recordD.Correlation3.Value IsNot Nothing Then
                returnMatrix(CInt(recordD.RowIndexNumber3.Value), CInt(recordD.ColumnIndexNumber3.Value)) = CDbl(recordD.Correlation3.Value)
            End If
            If recordD.RowIndexNumber4.Value IsNot Nothing And recordD.ColumnIndexNumber4.Value IsNot Nothing And recordD.Correlation4.Value IsNot Nothing Then
                returnMatrix(CInt(recordD.RowIndexNumber4.Value), CInt(recordD.ColumnIndexNumber4.Value)) = CDbl(recordD.Correlation4.Value)
            End If
            If recordD.RowIndexNumber5.Value IsNot Nothing And recordD.ColumnIndexNumber5.Value IsNot Nothing And recordD.Correlation5.Value IsNot Nothing Then
                returnMatrix(CInt(recordD.RowIndexNumber5.Value), CInt(recordD.ColumnIndexNumber5.Value)) = CDbl(recordD.Correlation5.Value)
            End If
        Next
        Return returnMatrix
    End Function

End Class