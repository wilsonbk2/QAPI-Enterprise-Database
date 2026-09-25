SELECT
    EmployeeOnlyCurrent.[Supervisor Name],
    EmployeeOnlyCurrent.[Full Name] AS Employee,
    EmployeeOnlyCurrent.JOBTITLE,
    EmployeeOnlyCurrent.[CC-BASE] AS [Cost Center],
    EmployeeOnlyCurrent.POD AS MainUnit,
    EmployeeOnlyCurrent.WARD AS SubUnit,
    EmployeeOnlyCurrent.NursingDesignationCode AS NursingStatus,
    EmployeeOnlyCurrent.[POSTYP DESC] AS PositionType

FROM
    EmployeeOnlyCurrent

ORDER BY
    EmployeeOnlyCurrent.[Supervisor Name]
    EmployeeOnlyCurrent.[Full Name];
