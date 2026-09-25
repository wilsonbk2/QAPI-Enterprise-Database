SELECT
    EmployeeOnlyCurrent.*,
    STAYINT.FORE,
    STAYINT.SDATE,
    STAYINT.STIME,
    STAYINT.NOTES,
    STAYFORE.FORECOLOR,
    STAYFORE.FOREDESC,
    DateSerial(
Year(Date()),
        1,
        1
    )AS CurrentYear,
    DateSerial(
Year(Date()),
        1,
        1
    )+364 AS Completeby,
    IIf(
        [STAYINT].[SDATE]>=[CurrentYear]
        AND
        [STAYINT].[SDATE]<=[Completeby],
        "Completed",
        IIf(
IsNull([STAYINT].[SDATE]),
            "Not Completed",
            IIf(
                [STAYINT].[SDATE]<[CurrentYear],
                "Expired",
                "Unicorn"
            )
        )
    )AS Status

FROM
    ((EmployeeOnlyCurrent LEFT JOIN StayInterviewsMostCurrent ON EmployeeOnlyCurrent.[EMP ID] = StayInterviewsMostCurrent.EMPID) LEFT JOIN STAYINT ON (StayInterviewsMostCurrent.SDATE = STAYINT.SDATE) AND (StayInterviewsMostCurrent.EMPID = STAYINT.EMPID))
    LEFT JOIN STAYFORE
        ON STAYINT.FORE = STAYFORE.FOREID;
