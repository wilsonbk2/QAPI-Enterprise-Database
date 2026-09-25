SELECT
    TrainingALLEmployeeHistory.EMPID,
    TrainingALLEmployeeHistory.[POSITION#],
    TrainingALLEmployeeHistory.[POS-START-DATE],
    TrainingEmpFirstESHsrtDate.[MinOfESH-HIRE-DATE] AS [FIRST-ESH-START-DATE],
    IIf(
IsNull([SEPARATION]),
        [POS-START-DATE]-1,
        IIf(
            [MinOfESH-HIRE-DATE]=[POS-START-DATE],
            [SEPARATION],
            [POS-START-DATE]-1
        )
    )AS [PREV-EXIT-DATE1],
    IIf(
IsNull([SEPARATION]),
Date(),
        [SEPARATION]
    )AS [POS-EXIT-DATE]

FROM
    TrainingALLEmployeeHistory
    INNER JOIN TrainingEmpFirstESHsrtDate
        ON TrainingALLEmployeeHistory.EMPID = TrainingEmpFirstESHsrtDate.EMPID

ORDER BY
    TrainingALLEmployeeHistory.SEPARATION DESC;
