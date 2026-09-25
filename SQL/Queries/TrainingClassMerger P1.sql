SELECT
    SDTCOURSECOMP.[Content Type],
    SDTCOURSECOMP.[Start Date],
    SDTCOURSECOMP.[Complete Date],
    SDTCOURSECOMP.[Last Launch],
    SDTCOURSECOMP.[Total Launches],
    SDTCOURSECOMP.[User ID],
    SDTCOURSECOMP.YearComp,
    [TempVars]![SetMergeName] AS Content,
    [TempVars]![SetIDName] AS ID,
    SDTCOURSECOMP.CCCID

FROM
    SDTCOURSECOMP

WHERE
    (((SDTCOURSECOMP.[Content Title])=[TempVars]![SetCourseName]));
