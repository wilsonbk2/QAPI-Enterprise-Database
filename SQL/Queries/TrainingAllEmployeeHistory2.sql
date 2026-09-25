SELECT
    TrainingALLEmployeeHistory.*,
    IIf(
IsNull([SEPARATION]),
Date(),
        [SEPARATION]
    )AS REALEXIT,
    TrainingEmpFirstESHsrtDate.[MinOfESH-HIRE-DATE] AS [A-ESH-HIRE-DATE]

FROM
    TrainingALLEmployeeHistory
    LEFT JOIN TrainingEmpFirstESHsrtDate
        ON TrainingALLEmployeeHistory.SSN = TrainingEmpFirstESHsrtDate.SSN;
