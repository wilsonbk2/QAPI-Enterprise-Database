SELECT
    EmployeeOnlyALL.[SUPV#],
    EmployeeOnlyALL.[EMP ID],
    EmployeeOnlyALL.[Full Name],
    EmployeeOnlyALL.JOBTITLE,
    EmployeeOnlyALL.[BLDG#],
    EmployeeOnlyALL.POD,
    EmployeeOnlyALL.WARD,
    IIf(
IsNull([APPROVAL]),
        "Send For Med Screen",
        [APPROVAL]
    )AS [APPROVAL??],
    FitTestingMostCurrentMask.EMPID AS FITEID,
    FitTestingMedApproval.EMPID AS MEDEID,
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
    )AS FINALAPPRVVV,
    FitTestingMedApproval.N95DATE,
    FitTestingMostCurrentMask.MaxOfFTDATE,
    EmployeeOnlyALL.[CC-BASE],
    EmployeeOnlyALL.[ROLE-CODE],
    EmployeeOnlyALL.NursingDesignationCode,
    IIf(
        [FINALAPPRV]="Schedule Fit Testing",
        "Mark Scheduled",
        ""
    )AS Scheduled,
    FitTestingMostCurrentSchedule.EMPID AS SCHID,
    IIf(
IsNull([SCHID])=False
        AND
IsNull([FITEID])=True,
        "Pending Fit Test",
        [FINALAPPRVVV]
    )AS FINALAPPRVOLD,
    IIf(
IsNull([ExpirationDate]),
        "No Mask Recorded",
        IIf(
            [ExpirationDate]<Date(),
            "Expired",
            "Current"
        )
    )AS MaskStatus,
    [MaxofFTDate]+365 AS ExpirationDate,
    IIf(
IsNull([MEDEID]),
        IIf(
IsNull([FITEID])=False,
            "Med Screen Form Missing",
            "Send For Med Screen"
        ),
        IIf(
            [APP]="NO",
            IIf(
IsNull([FITEID])=False
                AND
                [APP]="NO",
                "Screened and Fitted",
                "Denied for N95"
            ),
            IIf(
IsNull([SCHID]),
                IIf(
IsNull([FITEID]),
                    "Schedule Fit Testing",
                    "Screened and Fitted"
                ),
                IIf(
IsNull([FITEID]),
                    "Pending Fit Test",
                    "Screened and Fitted"
                )
            )
        )
    )AS FINALAPPRV,
    EmployeeOnlyALL.[Supervisor Name],
    EmployeeOnlyALL.SupvJobTitle,
    FitTestingMedApproval.APP

FROM
    ((EmployeeOnlyALL LEFT JOIN FitTestingMostCurrentSchedule ON EmployeeOnlyALL.[EMP ID] = FitTestingMostCurrentSchedule.EMPID) LEFT JOIN FitTestingMostCurrentMask ON EmployeeOnlyALL.[EMP ID] = FitTestingMostCurrentMask.EMPID)
    LEFT JOIN FitTestingMedApproval
        ON EmployeeOnlyALL.[EMP ID] = FitTestingMedApproval.EMPID

WHERE
    (((EmployeeOnlyALL.[SUPV#]) Is Not Null))

ORDER BY
    IIf(IsNull([APPROVAL]),"Send For Med Screen",[APPROVAL]);
