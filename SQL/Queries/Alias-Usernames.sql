SELECT
    tbl1Employees.Alias,
    [tbl1Employees].[FirstName] & " " & [tbl1Employees].[LastName] AS Fullname,
    SATELLITE2.[Full Name]

FROM
    tbl1Employees
    LEFT JOIN SATELLITE2
        ON tbl1Employees.EMPID = SATELLITE2.[EMP ID];
