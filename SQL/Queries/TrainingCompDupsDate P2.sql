SELECT
Max(SDTCOURSECOMP.CCCID) AS MaxOfCCCID,
    SDTCOURSECOMP.[User ID],
    TrainingCompDupsCrosstabDate.[User ID Field],
    TrainingCompDupsCrosstabDate.[COURSEID Field],
    TrainingCompDupsCrosstabDate.[YearComp Field],
    TrainingCompDupsCrosstabDate.NumberOfDups,
Min(SDTCOURSECOMP.[Last Launch]) AS [MinOfLast Launch],
    SDTCOURSECOMP.[Complete Date]

FROM
    SDTCOURSECOMP
    LEFT JOIN TrainingCompDupsCrosstabDate
        ON (SDTCOURSECOMP.COURSEID = TrainingCompDupsCrosstabDate.[COURSEID Field])
        AND
        (SDTCOURSECOMP.YearComp = TrainingCompDupsCrosstabDate.[YearComp Field])
        AND
        (SDTCOURSECOMP.[User ID] = TrainingCompDupsCrosstabDate.[User ID Field])

GROUP BY
    SDTCOURSECOMP.[User ID]
    TrainingCompDupsCrosstabDate.[User ID Field]
    TrainingCompDupsCrosstabDate.[COURSEID Field]
    TrainingCompDupsCrosstabDate.[YearComp Field]
    TrainingCompDupsCrosstabDate.NumberOfDups
    SDTCOURSECOMP.[Complete Date]

HAVING
    (((TrainingCompDupsCrosstabDate.[User ID Field]) Is Not Null));
