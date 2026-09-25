SELECT
    [INCIDENT Upload Table].[ID#],
    [INCIDENT Upload Table].Date,
    [INCIDENT Upload Table].Time,
    [INCIDENT Upload Table].Unit,
    [INCIDENT Upload Table].Location,
    [INCIDENT Upload Table].Injury,
    [INCIDENT Upload Table].Severity,
    [INCIDENT Upload Table].Incident,
    [INCIDENT Upload Table].Subcat,
    [INCIDENT Upload Table].Sitaution,
    [INCIDENT Upload Table].EventAggressor,
    [INCIDENT Upload Table].InjuredBodyPart

FROM
    [INCIDENT Upload Table]
    LEFT JOIN [INCIDENT OG Table]
        ON ([INCIDENT Upload Table].Subcat = [INCIDENT OG Table].Subcat)
        AND
        ([INCIDENT Upload Table].ActualTime = [INCIDENT OG Table].ActualTime)
        AND
        ([INCIDENT Upload Table].ActualDate = [INCIDENT OG Table].ActualDate)
        AND
        ([INCIDENT Upload Table].[ID#] = [INCIDENT OG Table].[ID#])

WHERE
    ((([INCIDENT OG Table].[ID#]) Is Null));
