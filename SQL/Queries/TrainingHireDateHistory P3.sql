SELECT
    [TrainingHireDateHistory P1].EMPID,
    [TrainingHireDateHistory P1].[ESH-HIRE-DATE],
    [TrainingHireDateHistory P1].[POS-START-DATE] AS [PREV-POS-START-DATE],
    [TrainingHireDateHistory P1].SEPARATION AS [PREV-EXIT-DATE],
    [TrainingHireDateHistory P2].[POS-START-DATE] AS [CURR-POS-START-DATE],
    [TrainingHireDateHistory P2].SEPARATION AS [CURR-EXIT-DATE],
    [TrainingHireDateHistory P1].EMPSTATUS AS [PREV-EMPSTATUS],
    [TrainingHireDateHistory P2].EMPSTATUS AS [CURR-EMPSTATUS]

FROM
    [TrainingHireDateHistory P1]
    INNER JOIN [TrainingHireDateHistory P2]
        ON [TrainingHireDateHistory P1].EMPID = [TrainingHireDateHistory P2].EMPID

ORDER BY
    [TrainingHireDateHistory P1].EMPID;
