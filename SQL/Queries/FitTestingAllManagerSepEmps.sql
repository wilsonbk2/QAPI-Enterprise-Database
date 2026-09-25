SELECT
    EmployeeOnlySeparated.[SUPV#],
    EmployeeOnlySeparated.[EMP ID],
    EmployeeOnlySeparated.[Full Name],
    EmployeeOnlySeparated.JOBTITLE,
    EmployeeOnlySeparated.[BLDG#],
    EmployeeOnlySeparated.POD,
    EmployeeOnlySeparated.WARD,
    IIf(
IsNull([APPROVAL]),
        "Send For Med Screen",
        [APPROVAL]
    )AS [APPROVAL??],
    FitTestingMostCurrentMask.EMPID AS FITEID,
    FitTestingMedApproval.EMPID AS MEDEID,
    FitTestingMedApproval.APPROVAL,
    IIf(
IsNull([MEDEID]),
        "Send For Med Screen",
        IIf(
IsNull([FITEID] And [MEDEID]),
            [APPROVAL??],
            [APPROVAL]
        )
    )AS APPRV,
    IIf(
IsNull([FITEID]),
        "Flag",
        "Noflag"
    )AS JFK,
    IIf(
        [APPRV]="Screened and Fitted"
        AND
        [JFK]="Flag",
        "Schedule Fit Testing",
        [APPRV]
    )AS FINALAPPRVV,
    IIf(
IsNull([FITEID])=False
        AND
IsNull([MEDEID]),
        "Med Screen Form Missing",
        [FINALAPPRVV]
    )AS FINALAPPRV

FROM
    (EmployeeOnlySeparated LEFT JOIN FitTestingMedApproval ON EmployeeOnlySeparated.[EMP ID] = FitTestingMedApproval.EMPID)
    LEFT JOIN FitTestingMostCurrentMask
        ON EmployeeOnlySeparated.[EMP ID] = FitTestingMostCurrentMask.EMPID

ORDER BY
    IIf(IsNull([APPROVAL]),"Send For Med Screen",[APPROVAL]);
