SELECT
    DISTINCTROW SDTCOURSECHECK.[Content Title],
    SDTCOURSECHECK.[Content Type],
    SDTCOURSECHECK.[Start Date],
    SDTCOURSECHECK.[Complete Date],
    SDTCOURSECHECK.[Last Launch],
    SDTCOURSECHECK.[Total Launches],
    SDTCOURSECHECK.[User ID],
    SDTCOURSECHECK.YearComp,
    [TrainingCompImport P2-5].COURSEID

FROM
    SDTCOURSECHECK
    LEFT JOIN [TrainingCompImport P2-5]
        ON SDTCOURSECHECK.[Content Title] = [TrainingCompImport P2-5].[Content Title];
