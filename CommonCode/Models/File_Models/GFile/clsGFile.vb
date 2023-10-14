'clsGFile
'Purpose:       To hold all the information from a GLOBAL POSITIONING SYSTEM DATA TRANSFER FORMAT
'               (G-FILE)
'Updates:
'Note:          
'               Public members & Functions PascalCase
'               private members & local variables camelCase
'Created on:    2023/08/04        
'By:            Grant Haynes

Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.IO

Public Class GFile

    Implements IWriteFile

    ' Properties
    Public Property RecordA As New GFileRecords.RecordA()
    Public Property RecordsB As New GFileRecordCollections.RecordsB()
    Public Property RecordsC As New GFileRecordCollections.RecordsC()
    Public Property RecordsD As New GFileRecordCollections.RecordsD()
    Public Property RecordsE As New GFileRecordCollections.RecordsE()
    Public Property RecordsF As New GFileRecordCollections.RecordsF()
    Public Property RecordsG As New GFileRecordCollections.RecordsG()
    Public Property RecordsH As New GFileRecordCollections.RecordsH()
    Public Property RecordsI As New GFileRecordCollections.RecordsI()

    'Default constructor
    Public Sub New()

    End Sub

    ' WriteFile, this writes an Bluebook file in a location
    Public Sub WriteFile(filepath As String) Implements IWriteFile.WriteFile

        Using file As New StreamWriter(filepath)

            file.WriteLine(RecordA.ToString())

            For Each recordB As GFileRecords.RecordB In RecordsB.Records
                file.WriteLine(recordB.ToString())

                For Each recordC As GFileRecords.RecordC In RecordsC.GetRecordsBySessionID(CInt(recordB.SessionID))
                    file.WriteLine(recordC.ToString())
                Next

                For Each recordD As GFileRecords.RecordD In RecordsD.GetRecordsBySessionID(CInt(recordB.SessionID))
                    file.WriteLine(recordD.ToString())
                Next

                For Each recordE As GFileRecords.RecordE In RecordsE.GetRecordsBySessionID(CInt(recordB.SessionID))
                    file.WriteLine(recordE.ToString())
                Next

                For Each recordF As GFileRecords.RecordF In RecordsF.GetRecordsBySessionID(CInt(recordB.SessionID))
                    file.WriteLine(recordF.ToString())
                Next

                For Each recordG As GFileRecords.RecordG In RecordsG.GetRecordsBySessionID(CInt(recordB.SessionID))
                    file.WriteLine(recordG.ToString())
                Next

                For Each recordH As GFileRecords.RecordH In RecordsH.GetRecordsBySessionID(CInt(recordB.SessionID))
                    file.WriteLine(recordH.ToString())
                Next

                For Each recordI As GFileRecords.RecordI In RecordsI.GetRecordsBySessionID(CInt(recordB.SessionID))
                    file.WriteLine(recordI.ToString())
                Next
            Next


        End Using
    End Sub

End Class