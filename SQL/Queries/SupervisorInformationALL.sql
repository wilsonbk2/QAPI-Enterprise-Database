SELECT
    SupervisorInformation.[EMP ID],
    SupervisorInformation.[Full Name],
    SATPOSITION.JOBTITLE,
    SATPOSITION.ESH_SubCat,
    ERCOSTCENTER.[CC-DESC],
    SupervisorPOSNumbers.[SUPV#],
    SATPOSITION.POD,
    SATPOSITION.WARD,
    SupervisorInformation.[EMP ID] AS SupervisorID,
    IIf(
IsNull([Full Name]),
        "VACANT",
        [Full Name]
    )AS [Supervisor Name]

FROM
    SupervisorPOSNumbers
    LEFT JOIN (ERCOSTCENTER RIGHT JOIN (SATPOSITION LEFT JOIN SupervisorInformation
        ON SATPOSITION.[POSITION#] = SupervisorInformation.[SUPV#]) ON ERCOSTCENTER.[CC-BASE] = SATPOSITION.[CC-BASE]) ON SupervisorPOSNumbers.[SUPV#] = SATPOSITION.[POSITION#]

WHERE
    (((SATPOSITION.JOBTITLE) Is Not Null));
