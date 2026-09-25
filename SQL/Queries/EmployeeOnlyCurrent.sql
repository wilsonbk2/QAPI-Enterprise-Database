SELECT
    SATTEMPKEY.SSN,
    SATELLITE2.[EMP ID],
    SATELLITE2.[Full Name],
    SATELLITE2.JOBTITLE,
    SATELLITE2.[BLDG#],
    SATELLITE2.POD,
    SATELLITE2.WARD,
    SATELLITE2.[SUPV#],
    SATELLITE2.[CC-BASE],
    SATELLITE2.[ROLE-CODE],
    SATELLITE2.NursingDesignationCode,
    SATELLITE2.[ESH-HIRE-DATE],
    SATELLITE2.[POS-START-DATE],
    SATELLITE2.SEPARATION,
    SATELLITE2.[POSTYP DESC],
    SATELLITE2.ESH_SubCat,
    SATELLITE2.[POSITION#],
    SATELLITE2.EMPSTATUS,
    SupervisorInformationALL.[Supervisor Name],
    SATELLITE2.FIRSTNAME,
    SATELLITE2.LASTNAME,
    SATELLITE2.MI,
    IIf(
IsNull([MaxOfSequentialKey]),
Val("1"),
        [MaxOfSequentialKey]+1
    )AS SequentialKey

FROM
    ((SATELLITE2 LEFT JOIN SupervisorInformationALL ON SATELLITE2.[SUPV#] = SupervisorInformationALL.[SUPV#]) LEFT JOIN SATTEMPKEY ON SATELLITE2.[EMP ID] = SATTEMPKEY.[EMP ID])
    LEFT JOIN TrainingEmpServiceMaxSequential
        ON SATTEMPKEY.SSN = TrainingEmpServiceMaxSequential.SSN

WHERE
    (((SATELLITE2.EMPSTATUS)="Current"));
