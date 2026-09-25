SELECT
    [SecRes Upload Table].[Patient ID],
    [SecRes Upload Table].SRType,
    [SecRes Upload Table].[Agency Name],
    [SecRes Upload Table].[Client Name],
    [SecRes Upload Table].[Type of Procedure Initied],
    [SecRes Upload Table].[Date/Time Procedure Initated],
    [SecRes Upload Table].[Date/Time Release],
    [SecRes Upload Table].[# of orders],
    [SecRes Upload Table].[Total Hours],
    [SecRes Upload Table].[Unit/Ward],
    [SecRes Upload Table].[Staff Init Proc],
    [SecRes Upload Table].[Mech Type Rest Description]

FROM
    [SecRes OG Table]
    RIGHT JOIN [SecRes Upload Table]
        ON ([SecRes OG Table].[Type of Procedure Initied] = [SecRes Upload Table].[Type of Procedure Initied])
        AND
        ([SecRes OG Table].ActualTime = [SecRes Upload Table].ActualTime)
        AND
        ([SecRes OG Table].ActualDate = [SecRes Upload Table].ActualDate)
        AND
        ([SecRes OG Table].[Patient ID] = [SecRes Upload Table].[Patient ID])

WHERE
    ((([SecRes OG Table].[Patient ID]) Is Null));
