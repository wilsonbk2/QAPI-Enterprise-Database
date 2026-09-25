SELECT
    VACCINE.EMPPAT,
    VACCINE.EID,
    SATELLITE2.[Full Name],
    VACCINES.VACKEY,
    VACCINES.VACDESC,
    VACCINES.VACPERIOD,
    VACCINE.VACDATE,
    [VACDATE]+[VACPERIOD] AS [Expiration Date],
    IIf(
Date()>[VACDATE]+[VACPERIOD],
        "Expired",
        "Current"
    )AS Status,
    SupervisorInformationALL.[Supervisor Name],
    SupervisorInformationALL.[CC-DESC],
    SATELLITE2.[CC-BASE],
    SATELLITE2.[BLDG#],
    SATELLITE2.WARD,
    "VaccineReport" AS Vaccine,
    SATELLITE2.EMPSTATUS

FROM
    ((VACCINE LEFT JOIN VACCINES ON VACCINE.VACCINE = VACCINES.VACKEY) LEFT JOIN SATELLITE2 ON VACCINE.EID = SATELLITE2.[EMP ID])
    LEFT JOIN SupervisorInformationALL
        ON SATELLITE2.[SUPV#] = SupervisorInformationALL.[SUPV#]

WHERE
    (((VACCINES.VACDESC) Like "*" & Forms!INFECTHandWashReports!EVENT & "*") And ((VACCINES.VACPERIOD)<>"0") And (([VACDATE]+[VACPERIOD]) Between Forms!INFECTHandWashReports!SDATE4 And Forms!INFECTHandWashReports!SDATE5) And ((SATELLITE2.EMPSTATUS)="Current"));
