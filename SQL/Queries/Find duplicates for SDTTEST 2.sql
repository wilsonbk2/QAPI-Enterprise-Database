SELECT
First([SDTTEST 2].[CCCID]) AS [CCCID Field],
Count([SDTTEST 2].[CCCID]) AS NumberOfDups

FROM
    [SDTTEST 2]

GROUP BY
    [SDTTEST 2].[CCCID]

HAVING
    (((Count([SDTTEST 2].[CCCID]))>1));
