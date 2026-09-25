SELECT
    EmployeeOnlyALL.[SUPV#],
    EmployeeOnlyALL.[EMP ID],
    EmployeeOnlyALL.[Full Name],
    EmployeeOnlyALL.JOBTITLE,
    FitTestingManagerDesignation.BRANDDESC,
    FitTestingManagerDesignation.MASKTYPE,
    FitTestingManagerDesignation.TRAINDESC,
    FitTestingManagerDesignation.DESC,
    EmployeeOnlyALL.[BLDG#] AS Building,
    EmployeeOnlyALL.POD AS MainU,
    EmployeeOnlyALL.WARD AS Unit,
    IIf(
IsNull([APPROVAL]),
        "Send For Med Screen",
        [APPROVAL]
    )AS [APPROVAL?],
    FitTestingManagerDesignation.FTDATE,
    [FTDATE]+365 AS ExpirationDate,
    IIf(
IsNull([ExpirationDate]),
        "No Mask Recorded",
        IIf(
            [ExpirationDate]<Date(),
            "Expired",
            "Current"
        )
    )AS MaskStatus,
    EmployeeOnlyALL.[Photo Path],
    EmployeeOnlyALL.[Supervisor Name],
    EmployeeOnlyALL.SupvJobTitle,
    [FitTestingAllManagerEmps Sep-Current].FINALAPPRV,
    EmployeeOnlyALL.NursingDesignationCode,
    [FitTestingAllManagerEmps Sep-Current].APP,
    IIf(
IsNull([N95DATE]),
        "Not Screened",
        [N95DATE]
    )AS Medscreenstatus,
    EmployeeOnlyALL.[CC-DESC],
    EmployeeOnlyALL.[CC-BASE],
    EmployeeOnlyALL.[POSTYP DESC],
    EmployeeOnlyALL.SEPARATION,
    EmployeeOnlyALL.EMPSTATUS,
    [Building] & "-" & [MainU] & "-" & [Unit] AS Location

FROM
    (FitTestingManagerDesignation RIGHT JOIN EmployeeOnlyALL ON FitTestingManagerDesignation.EMPID = EmployeeOnlyALL.[EMP ID])
    LEFT JOIN [FitTestingAllManagerEmps Sep-Current]
        ON FitTestingManagerDesignation.EMPID = [FitTestingAllManagerEmps Sep-Current].[EMP ID]

WHERE
    (((EmployeeOnlyALL.[SUPV#]) Is Not Null) And ((EmployeeOnlyALL.[EMP ID]) Like "*" & Forms!FitTestingReports!EVENT2 & "*"))

ORDER BY
    IIf(IsNull([APPROVAL]),"Send For Med Screen",[APPROVAL]);
