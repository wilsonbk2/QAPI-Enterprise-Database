SELECT
    TrainingAddNewYearP1.[ROLE-CODE],
    [MaxOfSDTYEAR]+1 AS SDTYEAR,
    DateSerial(
        [SDTYEAR],
        1,
        1
    )AS SDTDATEYEAR

FROM
    TrainingAddNewYearP1;
