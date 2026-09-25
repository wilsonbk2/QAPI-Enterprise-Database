SELECT
    tbl1Employees.EMPID,
    tbl1Employees.FirstName,
    tbl1Employees.LastName,
    tbl1Employees.Alias,
    tbl2EmployeeTypes.EmployeeType,
    tbl2EmployeeTypes.SECUREID,
    tbl5EmployeeSecurityTypes.SECURETYPE,
    tbl5EmployeeSecurityTypes.SECUREKEY

FROM
    tbl5EmployeeSecurityTypes
    INNER JOIN (tbl2EmployeeTypes INNER JOIN tbl1Employees
        ON tbl2EmployeeTypes.ID = tbl1Employees.EmployeeType_ID) ON tbl5EmployeeSecurityTypes.SECUREID = tbl2EmployeeTypes.SECUREID;
