SELECT
    EmployeeOnlyCurrent.[SUPV#],
    EmployeeOnlyCurrent.[EMP ID],
    EmployeeOnlyCurrent.[Full Name],
    EmployeeOnlyCurrent.JOBTITLE,
    EmployeeOnlyCurrent.[BLDG#],
    EmployeeOnlyCurrent.POD,
    EmployeeOnlyCurrent.WARD,
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
    )AS FINALAPPRVVV,
    FitTestingMedApproval.N95DATE,
    FitTestingMostCurrentMask.MaxOfFTDATE,
    EmployeeOnlyCurrent.[CC-BASE],
    EmployeeOnlyCurrent.[ROLE-CODE],
    EmployeeOnlyCurrent.NursingDesignationCode,
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
    SupervisorInformationALL.[Supervisor Name],
    IIf(
IsNull([MaxofFTDate]+365),
        "No Mask Recorded",
        IIf(
            [MaxofFTDate]+365<Date(),
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
    FitTestingAllMasks.MASKTYPE,
    FitTestingAllMasks.BRANDDESC,
    FitTestingAllMasks.SIZE,
    ERCOSTCENTER.[CC-DESC]

FROM
    ((SupervisorInformationALL RIGHT JOIN (FitTestingMostCurrentSchedule RIGHT JOIN (FitTestingMostCurrentMask RIGHT JOIN (FitTestingMedApproval RIGHT JOIN EmployeeOnlyCurrent ON FitTestingMedApproval.EMPID = EmployeeOnlyCurrent.[EMP ID]) ON FitTestingMostCurrentMask.EMPID = EmployeeOnlyCurrent.[EMP ID]) ON FitTestingMostCurrentSchedule.EMPID = EmployeeOnlyCurrent.[EMP ID]) ON SupervisorInformationALL.[SUPV#] = EmployeeOnlyCurrent.[SUPV#]) LEFT JOIN FitTestingAllMasks ON (FitTestingMostCurrentMask.MaxOfFTDATE = FitTestingAllMasks.FTDATE) AND (FitTestingMostCurrentMask.EMPID = FitTestingAllMasks.EMPID))
    LEFT JOIN ERCOSTCENTER
        ON EmployeeOnlyCurrent.[CC-BASE] = ERCOSTCENTER.[CC-BASE]

WHERE
    (((SupervisorInformationALL.[Supervisor Name]) Like "*" & Forms!FitTestingReports!EVENT & "*") And(
        ( IIf(
IsNull([MaxofFTDate]+365),
            "No Mask Recorded",
            IIf(
                [MaxofFTDate]+365<Date(),
                "Expired",
                "Current"
            )
        ) )="Expired"
    ) )

ORDER BY
    IIf(IsNull([APPROVAL]),"Send For Med Screen",[APPROVAL]);
