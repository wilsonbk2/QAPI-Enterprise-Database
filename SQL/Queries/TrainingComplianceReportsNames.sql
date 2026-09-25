SELECT
    [TrainingMostCurrentGLAIT P1].GLAIT,
    EmployeeOnlyALL.[Full Name]

FROM
    EmployeeOnlyALL
    RIGHT JOIN [TrainingMostCurrentGLAIT P1]
        ON EmployeeOnlyALL.SSN = [TrainingMostCurrentGLAIT P1].[MaxOfSSN]

GROUP BY
    [TrainingMostCurrentGLAIT P1].GLAIT
    EmployeeOnlyALL.[Full Name]
    EmployeeOnlyALL.[EMP ID]
    [TrainingMostCurrentGLAIT P1].MaxOfSSN
    EmployeeOnlyALL.SSN

ORDER BY
    EmployeeOnlyALL.[Full Name];
