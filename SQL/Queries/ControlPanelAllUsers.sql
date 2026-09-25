SELECT
    tbl1Employees.EMPID,
    tbl1Employees.FirstName,
    tbl1Employees.LastName,
    tbl1Employees.Password,
    tbl1Employees.DOB,
    tbl1Employees.EmployeeType_ID,
    tbl1Employees.Alias,
    tbl1Employees.Separation,
    tbl5EmployeeSecurityTypes.SECURETYPE

FROM
    tbl5EmployeeSecurityTypes
    INNER JOIN (tbl2EmployeeTypes INNER JOIN tbl1Employees
        ON tbl2EmployeeTypes.ID = tbl1Employees.EmployeeType_ID) ON tbl5EmployeeSecurityTypes.SECUREID = tbl2EmployeeTypes.SECUREID;
