SELECT
    EmployeeOnlyCurrent.[Full Name],
Val([EmployeeOnlyCurrent].[EMP ID])+Val([EmployeeOnlyCurrent].[POS-START-DATE]) AS Rand,
    EmployeeOnlyCurrent.[EMP ID],
    EmployeeOnlyCurrent.[POSTYP DESC]

FROM
    ((EmployeeOnlyCurrent LEFT JOIN StayInterviewsMostCurrent ON EmployeeOnlyCurrent.[EMP ID] = StayInterviewsMostCurrent.EMPID) LEFT JOIN STAYINT ON (StayInterviewsMostCurrent.SDATE = STAYINT.SDATE) AND (StayInterviewsMostCurrent.EMPID = STAYINT.EMPID))
    LEFT JOIN STAYFORE
        ON STAYINT.FORE = STAYFORE.FOREID

WHERE
    (((EmployeeOnlyCurrent.[Full Name]) Is Not Null) AND(
        ( IIf(
            [STAYINT].[SDATE]>=DateSerial(Year(Date()),1,1)
            AND
            [STAYINT].[SDATE]<=DateSerial(Year(Date()),1,1)+364,
            "Completed",
            IIf(
IsNull([STAYINT].[SDATE]),
                "Not Completed",
                IIf(
                    [STAYINT].[SDATE]<DateSerial(Year(Date()),1,1),
                    "Expired",
                    "Unicorn"
                )
            )
        ) )<>"Completed"
    ) AND ((EmployeeOnlyCurrent.[POSTYP DESC])="Classified" Or (EmployeeOnlyCurrent.[POSTYP DESC])="Hourly"));
