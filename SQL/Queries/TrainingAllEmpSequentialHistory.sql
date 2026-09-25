SELECT
    TrainingAllEmployeeHistory.SSN,
    TrainingAllEmployeeHistory.EMPID,
    TrainingAllEmployeeHistory.[POSITION#],
    IIf(
IsNull([MinOfESH-HIRE-DATE]),
        [TrainingAllEmployeeHistory].[ESH-HIRE-DATE],
        [MinOfESH-HIRE-DATE]
    )AS [FIRST-ESH-HIRE-DATE],
    IIf(
IsNull([STEP2-POS-START-DATE]),
        [POS-START-DATE],
        [STEP2-POS-START-DATE]
    )AS [PREV-POS-START-DATE],
    IIf(
IsNull([STEP1-EXIT-DATE]),
        IIf(
IsNull([SEPARATION]),
Date(),
            [SEPARATION]
        ),
        [STEP1-EXIT-DATE]
    )AS [PREV-EXIT-DATE],
    TrainingAllEmployeeHistory.[POS-START-DATE],
    IIf(
IsNull([SEPARATION]),
Date(),
        [SEPARATION]
    )AS [POS-EXIT-DATE],
    TrainingAllEmployeeHistory.SequentialKeyNum

FROM
    (TrainingAllEmployeeHistory LEFT JOIN TrainingEmpFirstESHDate ON TrainingAllEmployeeHistory.SSN = TrainingEmpFirstESHDate.SSN)
    LEFT JOIN TrainingEmpServiceSequentialKey
        ON (TrainingAllEmployeeHistory.SequentialKeyNum = TrainingEmpServiceSequentialKey.SequentialKeyNum)
        AND
        (TrainingAllEmployeeHistory.SSN = TrainingEmpServiceSequentialKey.SSN);
