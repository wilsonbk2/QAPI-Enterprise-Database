SELECT
    [Nursing Codes].DESC,
    IIf(
Date()=Date(),
        "Separated",
        "Something is seriously wrong xD"
    )AS EMPSTATUS,
    SupervisorInformationALL.[Supervisor Name],
    EMPSERVICE.[POSITION#],
    EMPSERVICE.[EMP ID],
    EMPSERVICE.[Full Name],
    EMPSERVICE.[POS-START-DATE],
    EMPSERVICE.[EXIT-DATE],
    EMPSERVICE.ROLE,
    EMPSERVICE.[JOB TITLE],
    EMPSERVICE.CC,
    EMPSERVICE.[CC-DESC],
    EMPSERVICE.[SUPV#],
    EMPSERVICE.BDG,
    EMPSERVICE.PD,
    EMPSERVICE.WARD,
    EMPSERVICE.[POSTYP DESC],
    EMPSERVICE.SUBCAT,
    EMPSERVICE.NursingDesignationCode,
    EMPSERVICE.[ESH-HIRE-DATE],
    EMPSERVICE.SequentialKey,
    SATTEMPKEY.SSN

FROM
    ((EMPSERVICE LEFT JOIN [Nursing Codes] ON EMPSERVICE.NursingDesignationCode = [Nursing Codes].Nursing_ID) LEFT JOIN SupervisorInformationALL ON EMPSERVICE.[SUPV#] = SupervisorInformationALL.[SUPV#])
    LEFT JOIN SATTEMPKEY
        ON EMPSERVICE.[EMP ID] = SATTEMPKEY.[EMP ID];
