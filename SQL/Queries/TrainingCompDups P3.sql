SELECT
    [TrainingCompDups P2].[User ID],
    [TrainingCompDups P1].[User ID],
    [TrainingCompDups P1].COURSEID,
    [TrainingCompDups P1].YearComp,
    [TrainingCompDups P1].[MinOfLast Launch],
    [TrainingCompDups P2].[MinOfLast Launch],
    [TrainingCompDups P2].[Complete Date],
    [TrainingCompDups P2].CCCID,
    [TrainingCompDups P2].[COURSEID Field],
    [TrainingCompDups P2].Duplicate

FROM
    [TrainingCompDups P2]
    LEFT JOIN [TrainingCompDups P1]
        ON ([TrainingCompDups P2].[MinOfLast Launch] = [TrainingCompDups P1].[MinOfLast Launch])
        AND
        ([TrainingCompDups P2].[YearComp Field] = [TrainingCompDups P1].YearComp)
        AND
        ([TrainingCompDups P2].[COURSEID Field] = [TrainingCompDups P1].COURSEID)
        AND
        ([TrainingCompDups P2].[User ID] = [TrainingCompDups P1].[User ID])

WHERE
    ((([TrainingCompDups P1].[User ID]) Is Null) AND (([TrainingCompDups P2].Duplicate) Is Null))

ORDER BY
    [TrainingCompDups P2].[User ID];
