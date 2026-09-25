SELECT
    (SELECT COUNT(*) FROM TrainingAllEmployeeHistory2 AS T2 WHERE T2.[EMPID] = T1.[EMPID] AND T2.[POS-START-DATE] <= T1.[POS-START-DATE]) AS RowNumber,
    T1.[EMPID] AS Expr1,
    T1.[POS-START-DATE]

FROM
    TrainingAllEmployeeHistory2 AS T1

ORDER BY
    T1.[EMPID]
    T1.[POS-START-DATE];
