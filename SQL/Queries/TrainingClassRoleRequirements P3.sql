SELECT
    [TrainingClassRoleRequirements P2].*,
    SupervisorInformationALL.[SUPV#] AS [CurrentSUPV#],
    SupervisorInformationALL.[Supervisor Name] AS [Current Supervisor],
    TrainingAllEmpSequentialHistory.[PREV-POS-START-DATE],
    TrainingAllEmpSequentialHistory.[PREV-EXIT-DATE],
    TrainingAllEmpSequentialHistory.[POS-EXIT-DATE],
    IIf(
        [DATEASSIGN]>=[PREV-POS-START-DATE]
        AND
        [DATEASSIGN]>[PREV-EXIT-DATE],
        "First Class Assignment",
        IIf(
            [PREV-POS-START-DATE]=[TrainingClassRoleRequirements P2].[POS-START-DATE]
            AND
            [PREV-EXIT-DATE]=[POS-EXIT-DATE],
            "First Class Assignment",
            "Already Assigned - Exempt"
        )
    )AS [FIRST-ASSIGN],
    SDTGLAIT.GLAIT,
    SATELLITE2.EMPSTATUS AS ActualEmpStatus,
    [TrainingClassRoleRequirements P2].EMPID

FROM
    (([TrainingClassRoleRequirements P2] LEFT JOIN (SATELLITE2 LEFT JOIN SupervisorInformationALL ON SATELLITE2.[SUPV#] = SupervisorInformationALL.[SUPV#]) ON [TrainingClassRoleRequirements P2].EMPID = SATELLITE2.[EMP ID]) LEFT JOIN SDTGLAIT ON [TrainingClassRoleRequirements P2].EMPID = SDTGLAIT.EMPID)
    INNER JOIN TrainingAllEmpSequentialHistory
        ON ([TrainingClassRoleRequirements P2].SSN = TrainingAllEmpSequentialHistory.SSN)
        AND
        ([TrainingClassRoleRequirements P2].[POS-START-DATE] = TrainingAllEmpSequentialHistory.[POS-START-DATE])

WHERE
    ((( IIf(
        [DATEASSIGN]>=[PREV-POS-START-DATE]
        AND
        [DATEASSIGN]>[PREV-EXIT-DATE],
        "First Class Assignment",
        IIf(
            [PREV-POS-START-DATE]=[TrainingClassRoleRequirements P2].[POS-START-DATE]
            AND
            [PREV-EXIT-DATE]=[POS-EXIT-DATE],
            "First Class Assignment",
            "Already Assigned - Exempt"
        )
    ) )<>"Already Assigned - Exempt"));
