SELECT
    TrainingEmpService.SSN,
Max(TrainingEmpService.SequentialKey) AS MaxOfSequentialKey

FROM
    TrainingEmpService

GROUP BY
    TrainingEmpService.SSN;
