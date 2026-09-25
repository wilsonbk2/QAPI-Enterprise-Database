SELECT
    EmployeeOnlyCurrent.[SUPV#],
    EmployeeOnlyCurrent.[EMP ID],
    EmployeeOnlyCurrent.[Full Name],
    EmployeeOnlyCurrent.JOBTITLE,
    FitTestingManagerDesignation.BRANDDESC,
    FitTestingManagerDesignation.MASKTYPE,
    FitTestingManagerDesignation.TRAINDESC,
    FitTestingManagerDesignation.DESC,
    EmployeeOnlyCurrent.[BLDG#],
    EmployeeOnlyCurrent.POD,
    EmployeeOnlyCurrent.WARD,
    IIf(
IsNull([APPROVAL]),
        "Send For Med Screen",
        [APPROVAL]
    )AS [APPROVAL?],
    FitTestingManagerDesignation.FTDATE,
    EmployeeOnlyCurrent.NursingDesignationCode

FROM
    FitTestingManagerDesignation
    RIGHT JOIN EmployeeOnlyCurrent
        ON FitTestingManagerDesignation.EMPID = EmployeeOnlyCurrent.[EMP ID]

ORDER BY
    IIf(IsNull([APPROVAL]),"Send For Med Screen",[APPROVAL]);
