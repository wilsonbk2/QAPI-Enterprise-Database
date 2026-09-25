SELECT
    VACCINE.VCKEY,
    VACCINE.EMPPAT,
    VACCINE.EID,
    SATELLITE2.JOBTITLE,
    VACCINES.VACDESC,
    VACCINE.VACDATE,
    VACCINE.SHIFT,
    VACCINE.NOTES,
    SATELLITE2.[Full Name],
    VACCINES.VACKEY,
    IIf(
        [VACKEY]="11"
        OR
        [VACKEY]="1"
        OR
        [VACKEY]="4",
        "Annual",
        IIf(
            [VACKEY]="10",
            "10 Year",
            "One-Time"
        )
    )AS Annual,
    IIf(
        [Annual]="One-Time",
        Null,
        DateSerial(
            [CurrentYear],
            [ESHMonth],
            [ESHDay]
        )
    )AS DueDate,
    IIf(
IsNull([DueDate]),
        Null,
        IIf(
IsNull([VACDATE]+365),
            "No Mask Recorded",
            IIf(
                [VACDATE]+365<Date(),
                "Expired",
                "Current"
            )
        )
    )AS VaccineStatus,
    VACCINE.UNIT,
    VACCINE.BLDG,
    UNIT.UNITMAIN,
    SATELLITE2.[ESH-HIRE-DATE],
Month([ESH-HIRE-DATE]) AS ESHMonth,
Year(Date()) AS CurrentYear,
Day([ESH-HIRE-DATE]) AS ESHDay,
    DateSerial(
        [CurrentYear],
        [ESHMonth],
        [ESHDay]
    )AS EXP,
    [DueDate]+365 AS RetestDate,
    [DueDate]-365 AS LastYearDue,
    SATELLITE2.[BLDG#],
    SATELLITE2.POD,
    SATELLITE2.WARD,
    BUILDING.[BUILDING#],
    SATELLITE2.EMPSTATUS,
    VACCINES.VACPERIOD,
    [VACDATE]+[VACPERIOD] AS ExpireDate,
    IIf(
Date()>[VACDATE]+[VACPERIOD],
        "Expired",
        IIf(
IsNull(Date()>[VACDATE]+[VACPERIOD]),
            "No Record",
            "Current"
        )
    )AS ExpireStatus,
    INFECTvaccinesubs.VACSUBDESC,
    VACCINE.OUTCOME,
    DEQRESOLUTION.DEQKEY

FROM
    (((((VACCINE LEFT JOIN VACCINES ON VACCINE.VACCINE = VACCINES.VACKEY) LEFT JOIN SATELLITE2 ON VACCINE.EID = SATELLITE2.[EMP ID]) LEFT JOIN UNIT ON VACCINE.UNIT = UNIT.UNITSUB) LEFT JOIN BUILDING ON SATELLITE2.[BLDG#] = BUILDING.BUI) LEFT JOIN INFECTvaccinesubs ON VACCINE.VACSUB = INFECTvaccinesubs.VACSUBKEY)
    LEFT JOIN DEQRESOLUTION
        ON VACCINE.VCKEY = DEQRESOLUTION.DEQKEY

WHERE
    (((DEQRESOLUTION.DEQKEY) Is Null))

ORDER BY
    VACCINE.VACDATE DESC;
