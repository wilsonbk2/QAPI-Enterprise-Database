SELECT
    TrainingMaxEmpServiceDate.EMPID,
    TrainingMaxEmpServiceDate.[POSITION#],
    TrainingMaxEmpServiceDate.[POS-EXIT-DATE],
    TrainingMaxEmpServiceDate.[FIRST-ESH-START-DATE],
    TrainingMaxEmpServiceDate.[POS-START-DATE],
    TrainingMaxEmpServiceDate.[POS-START-DATE],
    IIf(
IsNull([TrainingAllEmployeeHistory2].[POS-START-DATE]),
        [TrainingMaxEmpServiceDate].[POS-START-DATE],
        [TrainingAllEmployeeHistory2].[POS-START-DATE]
    )AS [PREV-POS-START-DATE],
    IIf(
IsNull([TrainingAllEmployeeHistory2].[POS-START-DATE]),
        [POS-EXIT-DATE],
        [PREV-EXIT-DATE1]
    )AS [PREV-EXIT-DATE]

FROM
    TrainingAllEmployeeHistory2
    RIGHT JOIN TrainingMaxEmpServiceDate
        ON (TrainingAllEmployeeHistory2.SEPARATION = TrainingMaxEmpServiceDate.[PREV-EXIT-DATE1])
        AND
        (TrainingAllEmployeeHistory2.EMPID = TrainingMaxEmpServiceDate.EMPID);
