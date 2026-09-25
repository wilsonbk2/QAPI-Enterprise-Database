SELECT
    SDTCOURSECOMP.COURSEID,
    SDTCOURSECOMP.[Content Title],
    SDTCOURSECOMP.[Last Launch],
    SDTCOURSECOMP.[User ID],
    SDTCOURSECOMP.YearComp

FROM
    SDTCOURSECOMP
    LEFT JOIN SDTCOURSECHECK
        ON (SDTCOURSECOMP.YearComp = SDTCOURSECHECK.YearComp)
        AND
        (SDTCOURSECOMP.[User ID] = SDTCOURSECHECK.[User ID])
        AND
        (SDTCOURSECOMP.[Last Launch] = SDTCOURSECHECK.[Last Launch])
        AND
        (SDTCOURSECOMP.[Content Title] = SDTCOURSECHECK.[Content Title])

WHERE
    (((SDTCOURSECOMP.[Content Title])=TempVars!SetCourseName) And ((SDTCOURSECHECK.[#]) Is Null));
