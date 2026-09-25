SELECT
    FitTestingAllManagerEmps.[SUPV#],
    FitTestingAllManagerEmps.[EMP ID],
    FitTestingAllManagerEmps.[Full Name],
    FitTestingAllManagerEmps.JOBTITLE,
    FitTestingAllManagerEmps.[BLDG#],
    FitTestingAllManagerEmps.POD,
    FitTestingAllManagerEmps.[APPROVAL??],
    FITTESTING.EMPID,
    IIf(
IsNull([EMPID]),
        "Schedule for Fit Test",
        [APPROVAL??]
    )AS [APPROVAL?]

FROM
    FitTestingAllManagerEmps
    LEFT JOIN FITTESTING
        ON FitTestingAllManagerEmps.[EMP ID] = FITTESTING.EMPID;
