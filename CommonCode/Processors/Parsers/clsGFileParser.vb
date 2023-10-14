'clsGFileParser
'Purpose:       To parse a GFile
'Updates:   
'Note: 
'               Public members & Functions PascalCase
'               private members & local variables camelCase
'Created on:    2023/08/04
'By:            Grant Haynes

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text.RegularExpressions

Public Class GFileParser

    'Properties
    Public Property ErrorFilePath As String = ""

    'Default constructor
    Public Sub New()

    End Sub

    Public Function ParseARecord(Line As String) As GFileRecords.RecordA
        Dim recordA As New GFileRecords.RecordA()
        recordA.JobCode.Value = recordA.JobCode.ParseandTrim(Line)
        recordA.ProjectStartDate.Value = recordA.ProjectStartDate.Parse(Line)
        recordA.ProjectEndDate.Value = recordA.ProjectEndDate.Parse(Line)
        recordA.ProjectTitle.Value = recordA.ProjectTitle.ParseandTrim(Line)
        Return recordA
    End Function

    Public Function ParseBRecord(Line As String, sessionID As Integer) As GFileRecords.RecordB
        Dim recordB As New GFileRecords.RecordB()
        recordB.ProjectStartDate.Value = recordB.ProjectStartDate.Parse(Line)
        recordB.ProjectEndDate.Value = recordB.ProjectEndDate.Parse(Line)
        recordB.NumberofVectorsintheSession.Value = recordB.NumberofVectorsintheSession.Parse(Line)
        recordB.SoftwareNameandVersion.Value = recordB.SoftwareNameandVersion.ParseandTrim(Line)
        recordB.OrbitSource.Value = recordB.OrbitSource.ParseandTrim(Line)
        recordB.OrbitAccuracyEstimate.Value = recordB.OrbitAccuracyEstimate.Parse(Line)
        recordB.SolutionCoordinateSystemCode.Value = recordB.SolutionCoordinateSystemCode.Parse(Line)
        recordB.SolutionMeteorologicalUseCode.Value = recordB.SolutionMeteorologicalUseCode.Parse(Line)
        recordB.SolutionIonosphereUseCode.Value = recordB.SolutionIonosphereUseCode.Parse(Line)
        recordB.SolutionTimeParameterUseCode.Value = recordB.SolutionTimeParameterUseCode.Parse(Line)
        recordB.VectorNominalAccuracyCode.Value = recordB.VectorNominalAccuracyCode.Parse(Line)
        recordB.ProcessingAgencyCode.Value = recordB.ProcessingAgencyCode.ParseandTrim(Line)
        recordB.ProcessingDate.Value = recordB.ProcessingDate.Parse(Line)
        recordB.SolutionType.Value = recordB.SolutionType.ParseandTrim(Line)
        recordB.ProjectID.Value = recordB.ProjectID.ParseandTrim(Line)
        recordB.SessionID = sessionID
        Return recordB
    End Function

    Public Function ParseCRecord(Line As String, sessionID As Integer) As GFileRecords.RecordC
        Dim recordC As New GFileRecords.RecordC()
        recordC.OriginSSN.Value = recordC.OriginSSN.Parse(Line)
        recordC.DifferentialSSN.Value = recordC.DifferentialSSN.Parse(Line)
        recordC.DeltaX.Value = recordC.DeltaX.Parse(Line)
        recordC.StandardDeviationX.Value = recordC.StandardDeviationX.Parse(Line)
        recordC.DeltaY.Value = recordC.DeltaY.Parse(Line)
        recordC.StandardDeviationY.Value = recordC.StandardDeviationY.Parse(Line)
        recordC.DeltaZ.Value = recordC.DeltaZ.Parse(Line)
        recordC.StandardDeviationZ.Value = recordC.StandardDeviationZ.Parse(Line)
        recordC.RejectionCode.Value = recordC.RejectionCode.ParseandTrim(Line)
        recordC.DataMediaIdentifierOriginStation.Value = recordC.DataMediaIdentifierOriginStation.ParseandTrim(Line)
        recordC.DataMediaIdentifierDifferentialStation.Value = recordC.DataMediaIdentifierDifferentialStation.ParseandTrim(Line)
        recordC.SessionID = sessionID
        Return recordC
    End Function

    Public Function ParseDRecord(Line As String, sessionID As Integer) As GFileRecords.RecordD
        Dim recordD As New GFileRecords.RecordD()
        recordD.RowIndexNumber1.Value = recordD.RowIndexNumber1.Parse(Line)
        recordD.ColumnIndexNumber1.Value = recordD.ColumnIndexNumber1.Parse(Line)
        recordD.Correlation1.Value = recordD.Correlation1.Parse(Line)
        recordD.RowIndexNumber2.Value = recordD.RowIndexNumber2.Parse(Line)
        recordD.ColumnIndexNumber2.Value = recordD.ColumnIndexNumber2.Parse(Line)
        recordD.Correlation2.Value = recordD.Correlation2.Parse(Line)
        recordD.RowIndexNumber3.Value = recordD.RowIndexNumber3.Parse(Line)
        recordD.ColumnIndexNumber3.Value = recordD.ColumnIndexNumber3.Parse(Line)
        recordD.Correlation3.Value = recordD.Correlation3.Parse(Line)
        recordD.RowIndexNumber4.Value = recordD.RowIndexNumber4.Parse(Line)
        recordD.ColumnIndexNumber4.Value = recordD.ColumnIndexNumber4.Parse(Line)
        recordD.Correlation4.Value = recordD.Correlation4.Parse(Line)
        recordD.RowIndexNumber5.Value = recordD.RowIndexNumber5.Parse(Line)
        recordD.ColumnIndexNumber5.Value = recordD.ColumnIndexNumber5.Parse(Line)
        recordD.Correlation5.Value = recordD.Correlation5.Parse(Line)
        recordD.SessionID = sessionID
        Return recordD
    End Function

    Public Function ParseERecord(Line As String, sessionID As Integer) As GFileRecords.RecordE
        Dim recordE As New GFileRecords.RecordE()
        recordE.RowIndexNumber1.Value = recordE.RowIndexNumber1.Parse(Line)
        recordE.ColumnIndexNumber1.Value = recordE.ColumnIndexNumber1.Parse(Line)
        recordE.Covariance1.Value = recordE.Covariance1.Parse(Line)
        recordE.RowIndexNumber2.Value = recordE.RowIndexNumber2.Parse(Line)
        recordE.ColumnIndexNumber2.Value = recordE.ColumnIndexNumber2.Parse(Line)
        recordE.Covariance2.Value = recordE.Covariance2.Parse(Line)
        recordE.RowIndexNumber3.Value = recordE.RowIndexNumber3.Parse(Line)
        recordE.ColumnIndexNumber3.Value = recordE.ColumnIndexNumber3.Parse(Line)
        recordE.Covariance3.Value = recordE.Covariance3.Parse(Line)
        recordE.RowIndexNumber4.Value = recordE.RowIndexNumber4.Parse(Line)
        recordE.ColumnIndexNumber4.Value = recordE.ColumnIndexNumber4.Parse(Line)
        recordE.Covariance4.Value = recordE.Covariance4.Parse(Line)
        recordE.SessionID = sessionID
        Return recordE
    End Function

    Public Function ParseFRecord(Line As String, sessionID As Integer) As GFileRecords.RecordF
        Dim recordF As New GFileRecords.RecordF()
        recordF.OriginStationSerialNumber.Value = recordF.OriginStationSerialNumber.Parse(Line)
        recordF.DifferentialStationSerialNumber.Value = recordF.DifferentialStationSerialNumber.Parse(Line)
        recordF.DeltaX.Value = recordF.DeltaX.Parse(Line)
        recordF.StandardDeviationX.Value = recordF.StandardDeviationX.Parse(Line)
        recordF.DeltaY.Value = recordF.DeltaY.Parse(Line)
        recordF.StandardDeviationY.Value = recordF.StandardDeviationY.Parse(Line)
        recordF.DeltaZ.Value = recordF.DeltaZ.Parse(Line)
        recordF.StandardDeviationZ.Value = recordF.StandardDeviationZ.Parse(Line)
        recordF.RejectionCode.Value = recordF.RejectionCode.ParseandTrim(Line)
        recordF.OriginStationManufacturerCode.Value = recordF.OriginStationManufacturerCode.ParseandTrim(Line)
        recordF.OriginStationUTCDayofYearofOccupation.Value = recordF.OriginStationUTCDayofYearofOccupation.Parse(Line)
        recordF.OriginStationYearofOccupation.Value = recordF.OriginStationYearofOccupation.Parse(Line)
        recordF.OriginStationSessionIndicator.Value = recordF.OriginStationSessionIndicator.ParseandTrim(Line)
        recordF.DifferentialStationManufacturerCode.Value = recordF.DifferentialStationManufacturerCode.ParseandTrim(Line)
        recordF.DifferentialStationDayofYear.Value = recordF.DifferentialStationDayofYear.Parse(Line)
        recordF.DifferentialStationYearofOccupation.Value = recordF.DifferentialStationYearofOccupation.Parse(Line)
        recordF.DifferentialStationSessionIndicator.Value = recordF.DifferentialStationSessionIndicator.ParseandTrim(Line)
        recordF.SessionID = sessionID
        Return recordF
    End Function

    Public Function ParseGRecord(Line As String, sessionID As Integer) As GFileRecords.RecordG
        Dim recordG As New GFileRecords.RecordG()
        recordG.RecordUsageCode.Value = recordG.RecordUsageCode.Parse(Line)
        recordG.StationSerialNumber.Value = recordG.StationSerialNumber.Parse(Line)
        recordG.ShortID.Value = recordG.ShortID.ParseandTrim(Line)
        recordG.CoordinateFrameDesignator.Value = recordG.CoordinateFrameDesignator.ParseandTrim(Line)
        recordG.Xcoordinate.Value = recordG.Xcoordinate.Parse(Line)
        recordG.Ycoordinate.Value = recordG.Ycoordinate.Parse(Line)
        recordG.Zcoordinate.Value = recordG.Zcoordinate.Parse(Line)
        recordG.SigmaX.Value = recordG.SigmaX.Parse(Line)
        recordG.SigmaY.Value = recordG.SigmaY.Parse(Line)
        recordG.SigmaZ.Value = recordG.SigmaZ.Parse(Line)
        recordG.SessionID = sessionID
        Return recordG
    End Function
    Public Function ParseHRecord(Line As String, sessionID As Integer) As GFileRecords.RecordH
        Dim recordH As New GFileRecords.RecordH()
        recordH.StationSerialNumber.Value = recordH.StationSerialNumber.Parse(Line)
        recordH.FourCharacterIdentifier.Value = recordH.FourCharacterIdentifier.ParseandTrim(Line)
        recordH.ExternalFrequencyStandardCode.Value = recordH.ExternalFrequencyStandardCode.Parse(Line)
        recordH.VectorMeteorologicalUseCode.Value = recordH.VectorMeteorologicalUseCode.Parse(Line)
        recordH.VectorTimeParameterUseCode.Value = recordH.VectorTimeParameterUseCode.Parse(Line)
        recordH.VectorIonosphereUseCode.Value = recordH.VectorIonosphereUseCode.Parse(Line)
        recordH.VectorSolutionType.Value = recordH.VectorSolutionType.ParseandTrim(Line)
        recordH.Comments.Value = recordH.Comments.ParseandTrim(Line)
        recordH.SessionID = sessionID
        Return recordH
    End Function

    Public Function ParseIRecord(Line As String, sessionID As Integer) As GFileRecords.RecordI
        Dim RecordI As New GFileRecords.RecordI()
        RecordI.AntennaPatternFileName.Value = RecordI.AntennaPatternFileName.ParseandTrim(Line)
        RecordI.AntennaPatternFileSource.Value = RecordI.AntennaPatternFileSource.ParseandTrim(Line)
        RecordI.AntennaPatternFileDate.Value = RecordI.AntennaPatternFileDate.Parse(Line)
        RecordI.SessionID = sessionID
        Return RecordI
    End Function

    Public Function ParseFile(GFilePath As String) As GFile
        Dim bluebookValueParser As New BluebookValueParser() 'Custom class to parse bluebook formatted floats
        Dim extension As String = Path.GetExtension(GFilePath)
        Me.ErrorFilePath = GFilePath.Replace(extension, "_parsing.err")
        Dim logger As New Logger(ErrorFilePath)
        Dim GFileReader As New IO.StreamReader(GFilePath)
        Dim GFile As New GFile()
        Dim sessionID As Integer = 0
        Try
            Do While GFileReader.Peek() <> -1
                Dim line As String = GFileReader.ReadLine()
                Dim recType As String = Mid(line, 1, 1)
                Select Case recType
                    Case "A"
                        GFile.RecordA = ParseARecord(line)
                    Case "B"
                        sessionID += 1
                        GFile.RecordsB.Records.Add(ParseBRecord(line, sessionID))
                    Case "C"
                        GFile.RecordsC.Records.Add(ParseCRecord(line, sessionID))
                    Case "D"
                        GFile.RecordsD.Records.Add(ParseDRecord(line, sessionID))
                    Case "E"
                        GFile.RecordsE.Records.Add(ParseERecord(line, sessionID))
                    Case "F"
                        GFile.RecordsF.Records.Add(ParseFRecord(line, sessionID))
                    Case "G"
                        GFile.RecordsG.Records.Add(ParseGRecord(line, sessionID))
                    Case "H"
                        GFile.RecordsH.Records.Add(ParseHRecord(line, sessionID))
                    Case "I"
                        GFile.RecordsI.Records.Add(ParseIRecord(line, sessionID))
                End Select
            Loop
            GFileReader.Close()
        Catch
            Throw New System.Exception("Error parsing horizontal bluebook file.")
        Finally
            GFileReader.Close()
        End Try
        Return GFile
    End Function

End Class