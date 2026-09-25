SELECT
    SATELLITE2.[EMP ID] AS EmployeeID,
    SATELLITE2.[SUPV#],
    SupervisorInformationALL.[Full Name],
    SupervisorInformationALL.SupervisorID,
    SATELLITE2.EMPSTATUS

FROM
    SupervisorInformationALL
    INNER JOIN SATELLITE2
        ON SupervisorInformationALL.[SUPV#] = SATELLITE2.[SUPV#]

WHERE
    (((SATELLITE2.EMPSTATUS)="current"));
