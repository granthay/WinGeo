'clsGFileRecords
'Purpose:       To hold all records from a G File
'Updates:
'Note: 
'               Public members & Functions PascalCase
'               private members & local variables camelCase
'Created on:    2023/08/04 
'By:            Grant Haynes

Option Strict On
Option Explicit On

Imports System.Collections.Generic

Public Class GFileRecords

    'RecordA
    'Purpose:       To hold all A record information, Project Record
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/08/04
    'By:            Grant Haynes

    Public Class RecordA

        Implements IToString

        Public Property JobCode As New FixedWidthPropertyClasses.StringProperty("", 2, 2, "[A-Z]")
        Public Property ProjectStartDate As New FixedWidthPropertyClasses.DateProperty(Nothing, 4, 8, "yyyyMMdd", New List(Of String)({"yyyyMMdd"}))
        Public Property ProjectEndDate As New FixedWidthPropertyClasses.DateProperty(Nothing, 12, 8, "yyyyMMdd", New List(Of String)({"yyyyMMdd"}))
        Public Property ProjectTitle As New FixedWidthPropertyClasses.StringProperty("", 20, 59, "[A-Z]")

        'Default Constructor
        Public Sub New()

        End Sub

        Public Overrides Function ToString() As String Implements IToString.ToString
            Dim format As String = "A{0}{1}{2}{3}"
            Dim returnString As String = String.Format(format,
                                                        JobCode.ToFixedWidthString(),
                                                        ProjectStartDate.ToFixedWidthString(),
                                                        ProjectEndDate.ToFixedWidthString(),
                                                        ProjectTitle.ToFixedWidthString())
            Return returnString.PadRight(80, " "c)
        End Function
    End Class

    'RecordB
    'Purpose:       To hold all B record information, Session Header Record
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/06/10
    'By:            Grant Haynes

    Public Class RecordB


        Implements IToString

        Public Property ProjectStartDate As New FixedWidthPropertyClasses.DateProperty(Nothing, 2, 12, "yyyyMMddHHmm", New List(Of String)({"yyyyMMddHHmm"}))
        Public Property ProjectEndDate As New FixedWidthPropertyClasses.DateProperty(Nothing, 14, 12, "yyyyMMddHHmm", New List(Of String)({"yyyyMMddHHmm"}))
        Public Property NumberofVectorsintheSession As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 26, 1)
        Public Property SoftwareNameandVersion As New FixedWidthPropertyClasses.StringProperty("", 28, 15, "[A-Z]")
        Public Property OrbitSource As New FixedWidthPropertyClasses.StringProperty("", 43, 5, "[A-Z]")
        Public Property OrbitAccuracyEstimate As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 48, 4, 2, 2)
        Public Property SolutionCoordinateSystemCode As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 52, 2)
        Public Property SolutionMeteorologicalUseCode As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 54, 2)
        Public Property SolutionIonosphereUseCode As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 56, 2)
        Public Property SolutionTimeParameterUseCode As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 58, 2)
        Public Property VectorNominalAccuracyCode As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 60, 1)
        Public Property ProcessingAgencyCode As New FixedWidthPropertyClasses.StringProperty("", 61, 6, "[A-Z]")
        Public Property ProcessingDate As New FixedWidthPropertyClasses.DateProperty(Nothing, 67, 8, "yyyyMMdd", New List(Of String)({"yyyyMMdd"}))
        Public Property SolutionType As New FixedWidthPropertyClasses.StringProperty("", 75, 5, "[A-Z]")
        Public Property ProjectID As New FixedWidthPropertyClasses.StringProperty("", 91, 104, "[A-Z]")
        Public Property SessionID As Integer? 'Since a file can have multiple Sessions, there needs to be an id to tie them together, Not parsed from file

        'Default Constructor
        Public Sub New()

        End Sub

        Public Overrides Function ToString() As String Implements IToString.ToString
            Dim format As String = "B{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}"
            Dim returnString As String = String.Format(format,
                                                        ProjectStartDate.ToFixedWidthString(),
                                                        ProjectEndDate.ToFixedWidthString(),
                                                        NumberofVectorsintheSession.ToFixedWidthString(),
                                                        SoftwareNameandVersion.ToFixedWidthString(),
                                                        OrbitSource.ToFixedWidthString(),
                                                        OrbitAccuracyEstimate.ToFixedWidthString(),
                                                        SolutionCoordinateSystemCode.ToFixedWidthString(),
                                                        SolutionMeteorologicalUseCode.ToFixedWidthString(),
                                                        SolutionIonosphereUseCode.ToFixedWidthString(),
                                                        SolutionTimeParameterUseCode.ToFixedWidthString(),
                                                        VectorNominalAccuracyCode.ToFixedWidthString(),
                                                        ProcessingAgencyCode.ToFixedWidthString(),
                                                        ProcessingDate.ToFixedWidthString(),
                                                        SolutionType.ToFixedWidthString(),
                                                        ProjectID.ToFixedWidthString())
            Return returnString.PadRight(80, " "c)
        End Function
    End Class

    'RecordC
    'Purpose:       To hold all C record information, Vector Record
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/06/10
    'By:            Grant Haynes

    Public Class RecordC

        Implements IToString

        Public Property OriginSSN As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 2, 4)
        Public Property DifferentialSSN As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 6, 4)
        Public Property DeltaX As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 10, 11, 7, 4)
        Public Property StandardDeviationX As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 21, 5, 1, 4)
        Public Property DeltaY As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 26, 11, 7, 4)
        Public Property StandardDeviationY As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 37, 5, 1, 4)
        Public Property DeltaZ As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 42, 11, 7, 4)
        Public Property StandardDeviationZ As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 53, 5, 1, 4)
        Public Property RejectionCode As New FixedWidthPropertyClasses.StringProperty("", 58, 1, "[A-Z]")
        Public Property DataMediaIdentifierOriginStation As New FixedWidthPropertyClasses.StringProperty("", 59, 10, "[A-Z]")
        Public Property DataMediaIdentifierDifferentialStation As New FixedWidthPropertyClasses.StringProperty("", 69, 10, "[A-Z]")
        Public Property SessionID As Integer 'Since a file can have multiple Sessions, there needs to be an id to tie them together, Not parsed from file

        'Default Constructor
        Public Sub New()

        End Sub

        Public Overrides Function ToString() As String Implements IToString.ToString
            Dim format As String = "C{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}"
            Dim returnString As String = String.Format(format,
                                                        OriginSSN.ToFixedWidthString(),
                                                        DifferentialSSN.ToFixedWidthString(),
                                                        DeltaX.ToFixedWidthString(),
                                                        StandardDeviationX.ToFixedWidthString(),
                                                        DeltaY.ToFixedWidthString(),
                                                        StandardDeviationY.ToFixedWidthString(),
                                                        DeltaZ.ToFixedWidthString(),
                                                        StandardDeviationZ.ToFixedWidthString(),
                                                        RejectionCode.ToFixedWidthString(),
                                                        DataMediaIdentifierOriginStation.ToFixedWidthString(),
                                                        DataMediaIdentifierDifferentialStation.ToFixedWidthString())
            Return returnString.PadRight(80, " "c)
        End Function
    End Class

    'RecordD
    'Purpose:       To hold all D record information, Correlation Record 
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/06/10
    'By:            Grant Haynes

    Public Class RecordD

        Implements IToString

        Public Property RowIndexNumber1 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 2, 3)
        Public Property ColumnIndexNumber1 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 5, 3)
        Public Property Correlation1 As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 8, 9, 2, 7)
        Public Property RowIndexNumber2 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 17, 3)
        Public Property ColumnIndexNumber2 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 20, 3)
        Public Property Correlation2 As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 23, 9, 2, 7)
        Public Property RowIndexNumber3 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 32, 3)
        Public Property ColumnIndexNumber3 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 35, 3)
        Public Property Correlation3 As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 38, 9, 2, 7)
        Public Property RowIndexNumber4 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 47, 3)
        Public Property ColumnIndexNumber4 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 50, 3)
        Public Property Correlation4 As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 53, 9, 2, 7)
        Public Property RowIndexNumber5 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 62, 3)
        Public Property ColumnIndexNumber5 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 65, 3)
        Public Property Correlation5 As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 68, 9, 2, 7)
        Public Property SessionID As Integer 'Since a file can have multiple Sessions, there needs to be an id to tie them together, Not parsed from file

        'Default Constructor
        Public Sub New()

        End Sub

        Public Overrides Function ToString() As String Implements IToString.ToString
            Dim format As String = "D{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}"
            Dim returnString As String = String.Format(format,
                                                        RowIndexNumber1.ToFixedWidthString(),
                                                        ColumnIndexNumber1.ToFixedWidthString(),
                                                        Correlation1.ToFixedWidthString(),
                                                        RowIndexNumber2.ToFixedWidthString(),
                                                        ColumnIndexNumber2.ToFixedWidthString(),
                                                        Correlation2.ToFixedWidthString(),
                                                        RowIndexNumber3.ToFixedWidthString(),
                                                        ColumnIndexNumber3.ToFixedWidthString(),
                                                        Correlation3.ToFixedWidthString(),
                                                        RowIndexNumber4.ToFixedWidthString(),
                                                        ColumnIndexNumber4.ToFixedWidthString(),
                                                        Correlation4.ToFixedWidthString(),
                                                        RowIndexNumber5.ToFixedWidthString(),
                                                        ColumnIndexNumber5.ToFixedWidthString(),
                                                        Correlation5.ToFixedWidthString())
            Return returnString.PadRight(80, " "c)
        End Function
    End Class

    'RecordE
    'Purpose:       To hold all E record information, Covariance Record
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/06/10
    'By:            Grant Haynes

    Public Class RecordE

        Implements IToString

        Public Property RowIndexNumber1 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 2, 3)
        Public Property ColumnIndexNumber1 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 5, 3)
        Public Property Covariance1 As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 8, 12, 4, 8)
        Public Property RowIndexNumber2 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 20, 3)
        Public Property ColumnIndexNumber2 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 23, 3)
        Public Property Covariance2 As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 26, 12, 4, 8)
        Public Property RowIndexNumber3 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 38, 3)
        Public Property ColumnIndexNumber3 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 41, 3)
        Public Property Covariance3 As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 44, 12, 4, 8)
        Public Property RowIndexNumber4 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 56, 3)
        Public Property ColumnIndexNumber4 As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 59, 3)
        Public Property Covariance4 As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 62, 12, 4, 8)
        Public Property SessionID As Integer 'Since a file can have multiple Sessions, there needs to be an id to tie them together, Not parsed from file

        'Default Constructor
        Public Sub New()

        End Sub

        Public Overrides Function ToString() As String Implements IToString.ToString
            Dim format As String = "E{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}"
            Dim returnString As String = String.Format(format,
                                                        RowIndexNumber1.ToFixedWidthString(),
                                                        ColumnIndexNumber1.ToFixedWidthString(),
                                                        Covariance1.ToFixedWidthString(),
                                                        RowIndexNumber2.ToFixedWidthString(),
                                                        ColumnIndexNumber2.ToFixedWidthString(),
                                                        Covariance2.ToFixedWidthString(),
                                                        RowIndexNumber3.ToFixedWidthString(),
                                                        ColumnIndexNumber3.ToFixedWidthString(),
                                                        Covariance3.ToFixedWidthString(),
                                                        RowIndexNumber4.ToFixedWidthString(),
                                                        ColumnIndexNumber4.ToFixedWidthString(),
                                                        Covariance4.ToFixedWidthString())
            Return returnString.PadRight(80, " "c)
        End Function
    End Class

    'RecordF
    'Purpose:       To hold all F record information, Long Vector Record
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/06/10
    'By:            Grant Haynes

    Public Class RecordF

        Implements IToString

        Public Property OriginStationSerialNumber As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 2, 4)           '(ssn) (vector tail)
        Public Property DifferentialStationSerialNumber As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 6, 4)     '(vector head)
        Public Property DeltaX As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 10, 13, 9, 4)                             '(XXXXXXXXX.xxxx meters) 
        Public Property StandardDeviationX As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 23, 5, 1, 4)                  '(X.xxxx meters) 
        Public Property DeltaY As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 28, 13, 9, 4)                             '(XXXXXXXXX.xxxx meters) 
        Public Property StandardDeviationY As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 41, 5, 1, 4)                  '(X.xxxx meters)
        Public Property DeltaZ As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 46, 13, 9, 4)                             '(XXXXXXXXX.xxxx meters)
        Public Property StandardDeviationZ As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 59, 15, 1, 4)                 '(X.xxxx meters)
        Public Property RejectionCode As New FixedWidthPropertyClasses.StringProperty("", 64, 1, "[A-Z]")                   '(use upper case R to reject)
        Public Property OriginStationManufacturerCode As New FixedWidthPropertyClasses.StringProperty("", 65, 1, "[A-Z]")   '(N-6)
        Public Property OriginStationUTCDayofYearofOccupation As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 66, 3)  '(DDD)
        Public Property OriginStationYearofOccupation As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 69, 1)      '(Y) UTC
        Public Property OriginStationSessionIndicator As New FixedWidthPropertyClasses.StringProperty("", 70, 1, "[A-Z]")
        Public Property DifferentialStationManufacturerCode As New FixedWidthPropertyClasses.StringProperty("", 71, 1, "[A-Z]") '(N-6)
        Public Property DifferentialStationDayofYear As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 72, 3)           '(DDD) UTC
        Public Property DifferentialStationYearofOccupation As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 75, 1)    '(Y) UTC
        Public Property DifferentialStationSessionIndicator As New FixedWidthPropertyClasses.StringProperty("", 76, 1, "[A-Z]")
        Public Property SessionID As Integer 'Since a file can have multiple Sessions, there needs to be an id to tie them together, Not parsed from file

        'Default Constructor
        Public Sub New()

        End Sub

        Public Overrides Function ToString() As String Implements IToString.ToString
            Dim format As String = "F{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}"
            Dim returnString As String = String.Format(format,
                                                        OriginStationSerialNumber.ToFixedWidthString(),
                                                        DifferentialStationSerialNumber.ToFixedWidthString(),
                                                        DeltaX.ToFixedWidthString(),
                                                        StandardDeviationX.ToFixedWidthString(),
                                                        DeltaY.ToFixedWidthString(),
                                                        StandardDeviationY.ToFixedWidthString(),
                                                        DeltaZ.ToFixedWidthString(),
                                                        StandardDeviationZ.ToFixedWidthString(),
                                                        RejectionCode.ToFixedWidthString(),
                                                        OriginStationManufacturerCode.ToFixedWidthString(),
                                                        OriginStationUTCDayofYearofOccupation.ToFixedWidthString(),
                                                        OriginStationYearofOccupation.ToFixedWidthString(),
                                                        OriginStationSessionIndicator.ToFixedWidthString(),
                                                        DifferentialStationManufacturerCode.ToFixedWidthString(),
                                                        DifferentialStationDayofYear.ToFixedWidthString(),
                                                        DifferentialStationYearofOccupation.ToFixedWidthString(),
                                                        DifferentialStationSessionIndicator.ToFixedWidthString())
            Return returnString.PadRight(80, " "c)
        End Function
    End Class

    'RecordG
    'Purpose:       To hold all G record information, Coordinate Record
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/06/10
    'By:            Grant Haynes

    Public Class RecordG

        Implements IToString

        Public Property RecordUsageCode As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 3, 1)
        Public Property StationSerialNumber As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 6, 4)
        Public Property ShortID As New FixedWidthPropertyClasses.StringProperty("", 11, 5, "[A-Z0-9]")                      '"short" station name
        Public Property CoordinateFrameDesignator As New FixedWidthPropertyClasses.StringProperty("", 16, 6, "[A-Z0-9]")    '(e.g. NAD 83, WGS 84, NAD 27,WGS 72, ITR 90, etc.; inquire for additions)
        Public Property Xcoordinate As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 22, 12, 8, 4)       '(XXXXXXXX.xxxx meters)
        Public Property Ycoordinate As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 35, 12, 8, 4)       '(YYYYYYYY.yyyy meters)
        Public Property Zcoordinate As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 48, 12, 8, 4)       '(ZZZZZZZZ.zzzz meters)
        Public Property SigmaX As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 61, 4, 2, 2)             '(SS.ss cm) blank if unknown Or greater than 99.99 cm
        Public Property SigmaY As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 56, 4, 2, 2)             '(SS.ss cm) blank if unknown Or greater than 99.99 cm
        Public Property SigmaZ As New FixedWidthPropertyClasses.DecimallessDoubleProperty(Nothing, 71, 4, 2, 2)             '(SS.ss cm) blank if unknown Or greater than 99.99 cm
        Public Property SessionID As Integer 'Since a file can have multiple Sessions, there needs to be an id to tie them together, Not parsed from file

        'Default Constructor
        Public Sub New()

        End Sub

        Public Overrides Function ToString() As String Implements IToString.ToString
            Dim format As String = "G {0} {1} {2} {3} {4} {5} {6} {7} {8} {9}"
            Dim returnString As String = String.Format(format,
                                                        RecordUsageCode.ToFixedWidthString(),
                                                        StationSerialNumber.ToFixedWidthString(),
                                                        ShortID.ToFixedWidthString(),
                                                        CoordinateFrameDesignator.ToFixedWidthString(),
                                                        Xcoordinate.ToFixedWidthString(),
                                                        Ycoordinate.ToFixedWidthString(),
                                                        Zcoordinate.ToFixedWidthString(),
                                                        SigmaX.ToFixedWidthString(),
                                                        SigmaY.ToFixedWidthString(),
                                                        SigmaZ.ToFixedWidthString())
            Return returnString.PadRight(80, " "c)
        End Function
    End Class

    'RecordH
    'Purpose:       To hold all H record information
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/06/10
    'By:            Grant Haynes

    Public Class RecordH

        Implements IToString

        Public Property StationSerialNumber As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 2, 4)     '(ssn)
        Public Property FourCharacterIdentifier As New FixedWidthPropertyClasses.StringProperty("", 6, 4, "[A-Z0-9]")
        Public Property ExternalFrequencyStandardCode As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 10, 2)  '(see table)
        Public Property VectorMeteorologicalUseCode As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 12, 2)    '(see table)
        Public Property VectorTimeParameterUseCode As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 14, 2)     '(see table)
        Public Property VectorIonosphereUseCode As New FixedWidthPropertyClasses.IntegerProperty(Nothing, 16, 2)        '(see table)
        Public Property VectorSolutionType As New FixedWidthPropertyClasses.StringProperty("", 18, 6, "[A-Z0-9]")       '(see table)
        Public Property Comments As New FixedWidthPropertyClasses.StringProperty("", 24, 55, "[A-Z0-9]")
        Public Property SessionID As Integer 'Since a file can have multiple Sessions, there needs to be an id to tie them together, Not parsed from file

        'Default Constructor
        Public Sub New()

        End Sub

        Public Overrides Function ToString() As String Implements IToString.ToString
            Dim format As String = "H{0}{1}{2}{3}{4}{5}{6}{7}"
            Dim returnString As String = String.Format(format,
                                                        StationSerialNumber.ToFixedWidthString(),
                                                        FourCharacterIdentifier.ToFixedWidthString(),
                                                        ExternalFrequencyStandardCode.ToFixedWidthString(),
                                                        VectorMeteorologicalUseCode.ToFixedWidthString(),
                                                        VectorTimeParameterUseCode.ToFixedWidthString(),
                                                        VectorIonosphereUseCode.ToFixedWidthString(),
                                                        VectorSolutionType.ToFixedWidthString(),
                                                        Comments.ToFixedWidthString())
            Return returnString.PadRight(80, " "c)
        End Function
    End Class

    'RecordI
    'Purpose:       To hold all I record information, Session Model Record
    'Updates:
    'Note: 
    '               Public members & Functions PascalCase
    '               private members & local variables camelCase
    'Created on:    2022/06/10
    'By:            Grant Haynes

    Public Class RecordI

        Implements IToString

        Public Property AntennaPatternFileName As New FixedWidthPropertyClasses.StringProperty("", 2, 20, "[A-Z0-9]")
        Public Property AntennaPatternFileSource As New FixedWidthPropertyClasses.StringProperty("", 22, 6, "[A-Z0-9]")
        Public Property AntennaPatternFileDate As New FixedWidthPropertyClasses.DateProperty(Nothing, 28, 8, "yyyyMMdd", New List(Of String)({"yyyyMMdd"}))
        Public Property SessionID As Integer 'Since a file can have multiple Sessions, there needs to be an id to tie them together, Not parsed from file

        'Default Constructor
        Public Sub New()

        End Sub

        Public Overrides Function ToString() As String Implements IToString.ToString
            Dim format As String = "I{0}{1}{2}"
            Dim returnString As String = String.Format(format,
                                                        AntennaPatternFileName.ToFixedWidthString(),
                                                        AntennaPatternFileSource.ToFixedWidthString(),
                                                        AntennaPatternFileDate.ToFixedWidthString())
            Return returnString.PadRight(80, " "c)
        End Function
    End Class

End Class