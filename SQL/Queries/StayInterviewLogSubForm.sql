SELECT
    STAYINT.SIID,
    STAYINT.EMPID,
    STAYFORE.FORECOLOR,
    STAYFORE.FOREDESC,
    STAYINT.SDATE,
    STAYINT.STIME,
    STAYINT.NOTES,
    EmployeeOnlyALL.[BLDG#],
    EmployeeOnlyALL.POD,
    EmployeeOnlyALL.WARD,
    EmployeeOnlyALL.EMPSTATUS,
    STAYFORE.FOREID,
    IIf(
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
    )AS Status

FROM
    ((STAYINT LEFT JOIN STAYFORE ON STAYINT.FORE = STAYFORE.FOREID) LEFT JOIN EmployeeOnlyALL ON STAYINT.EMPID = EmployeeOnlyALL.[EMP ID])
    LEFT JOIN DEQRESOLUTION
        ON STAYINT.SIKEY = DEQRESOLUTION.DEQKEY

WHERE
    (((DEQRESOLUTION.DEQKEY) Is Null));
