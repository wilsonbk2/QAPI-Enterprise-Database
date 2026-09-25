SELECT
    SDTCOURSECOMP.CCCID,
    SDTCOURSECOMP.[User ID],
    TrainingCompDupsCrosstab.[User ID Field],
    TrainingCompDupsCrosstab.[COURSEID Field],
    TrainingCompDupsCrosstab.[YearComp Field],
    TrainingCompDupsCrosstab.NumberOfDups,
Min(SDTCOURSECOMP.[Last Launch]) AS [MinOfLast Launch],
    SDTCOURSECOMP.[Complete Date],
    SDTCOURSECOMP.Duplicate

FROM
    SDTCOURSECOMP
    LEFT JOIN TrainingCompDupsCrosstab
        ON (SDTCOURSECOMP.YearComp = TrainingCompDupsCrosstab.[YearComp Field])
        AND
        (SDTCOURSECOMP.COURSEID = TrainingCompDupsCrosstab.[COURSEID Field])
        AND
        (SDTCOURSECOMP.[User ID] = TrainingCompDupsCrosstab.[User ID Field])

GROUP BY
    SDTCOURSECOMP.CCCID
    SDTCOURSECOMP.[User ID]
    TrainingCompDupsCrosstab.[User ID Field]
    TrainingCompDupsCrosstab.[COURSEID Field]
    TrainingCompDupsCrosstab.[YearComp Field]
    TrainingCompDupsCrosstab.NumberOfDups
    SDTCOURSECOMP.[Complete Date]
    SDTCOURSECOMP.Duplicate

HAVING
    (((TrainingCompDupsCrosstab.[User ID Field]) Is Not Null));
