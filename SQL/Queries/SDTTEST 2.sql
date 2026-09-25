SELECT
    SDTCOURSECOMP.CCCID,
    SDTCOURSECOMP.[Content Title],
    SDTCOURSECOMP.[User ID],
    SDTCOURSECOMP.COURSEID,
    [SDTTEST 1].COURSEID,
    [SDTTEST 1].YearComp,
    [SDTTEST 1].[User ID],
    [SDTTEST 1].[MaxOfLast Launch]

FROM
    SDTCOURSECOMP
    INNER JOIN [SDTTEST 1]
        ON (SDTCOURSECOMP.[User ID] = [SDTTEST 1].[User ID])
        AND
        (SDTCOURSECOMP.YearComp = [SDTTEST 1].YearComp)
        AND
        (SDTCOURSECOMP.COURSEID = [SDTTEST 1].COURSEID)
        AND
        (SDTCOURSECOMP.[Last Launch] = [SDTTEST 1].[MaxOfLast Launch])

WHERE
    (((SDTCOURSECOMP.COURSEID) Is Not Null) AND (([SDTTEST 1].[MaxOfLast Launch]) Is Not Null))

ORDER BY
    SDTCOURSECOMP.[User ID];
