SELECT
    TrainingALLEmployeeHistory.EMPID,
Min(TrainingALLEmployeeHistory.[ESH-HIRE-DATE]) AS [MinOfESH-HIRE-DATE],
    TrainingALLEmployeeHistory.SSN

FROM
    TrainingALLEmployeeHistory

GROUP BY
    TrainingALLEmployeeHistory.EMPID
    TrainingALLEmployeeHistory.SSN;
