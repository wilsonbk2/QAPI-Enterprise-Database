SELECT
    EmployeeIDList.[EMP ID],
    EmployeeOnlyALL.[Full Name],
    EmployeeOnlyALL.EMPSTATUS,
    IIf(
        [EMPSTATUS]="Current",
        [EmployeeOnlyAll].[JOBTITLE],
        [EMPSERVICE].[JOB TITLE]
    )AS JOBTITLEfin,
    IIf(
        [EMPSTATUS]="Current",
        [EmployeeOnlyAll].[ROLE-CODE],
        [EMPSERVICE].[ROLE]
    )AS ROLEfin,
    IIf(
        [EMPSTATUS]="Current",
        [EmployeeOnlyAll].[CC-BASE],
        [EMPSERVICE].[CC]
    )AS COSTCENTERfin,
    IIf(
        [EMPSTATUS]="Current",
        [EmployeeOnlyAll].[BLDG#],
        [EMPSERVICE].[BDG]
    )AS BDGfin,
    IIf(
        [EMPSTATUS]="Current",
        [EmployeeOnlyAll].[POD],
        [EMPSERVICE].[PD]
    )AS PODfin,
    IIf(
        [EMPSTATUS]="Current",
        [EmployeeOnlyAll].[WARD],
        [EMPSERVICE].[WARD]
    )AS WARDfin,
    IIf(
        [EMPSTATUS]="Current",
        [EmployeeOnlyAll].[POS-START-DATE],
        [EMPSERVICE].[POS-START-DATE]
    )AS [POS-START-DATEfin],
    IIf(
        [EMPSTATUS]="Current",
Date(),
        [EMPSERVICE].[EXIT-DATE]
    )AS CurrentDatefin,
    EmployeeOnlyALL.[SUPV#]

FROM
    (EmployeeIDList LEFT JOIN EmployeeOnlyALL ON EmployeeIDList.[EMP ID] = EmployeeOnlyALL.[EMP ID])
    LEFT JOIN EMPSERVICE
        ON EmployeeIDList.[EMP ID] = EMPSERVICE.[EMP ID]

WHERE
    (((EmployeeOnlyALL.[SUPV#]) Is Not Null));
