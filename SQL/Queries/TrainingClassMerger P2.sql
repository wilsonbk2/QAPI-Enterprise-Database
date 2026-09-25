SELECT
    SDTCOURSECOMP.[Content Title],
    SDTCOURSECOMP.[Last Launch],
    SDTCOURSECOMP.[User ID],
    SDTCOURSECOMP.YearComp

FROM
    SDTCOURSECOMP

WHERE
    (((SDTCOURSECOMP.[Content Title])=[TempVars]![SetMergeName]));
