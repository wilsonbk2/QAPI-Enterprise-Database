SELECT
    tbl1Employees.EMPID,
    tbl1Employees.FirstName,
    tbl1Employees.LastName,
    tbl1Employees.Password,
    tbl1Employees.DOB,
    tbl1Employees.EmployeeType_ID,
    tbl1Employees.Alias,
    tbl1Employees.Separation,
    tbl1Employees.DatabaseNumber,
    tbl4UserAlias.UserName,
    tbl4UserAlias.UserAlias,
    tbl1Employees.EMPID

FROM
    tbl1Employees
    INNER JOIN tbl4UserAlias
        ON tbl1Employees.Alias = tbl4UserAlias.Alias

GROUP BY
    tbl1Employees.FirstName
    tbl1Employees.LastName
    tbl1Employees.Password
    tbl1Employees.DOB
    tbl1Employees.EmployeeType_ID
    tbl1Employees.Alias
    tbl1Employees.Separation
    tbl1Employees.DatabaseNumber
    tbl4UserAlias.UserName
    tbl4UserAlias.UserAlias
    tbl1Employees.EMPID;
