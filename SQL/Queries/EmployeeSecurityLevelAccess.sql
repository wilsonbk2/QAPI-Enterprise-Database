SELECT
    tbl1Employees.FirstName,
    tbl1Employees.LastName,
    tbl2EmployeeTypes.EmployeeType,
    tbl5EmployeeSecurityTypes.SECUREID,
    tbl5EmployeeSecurityTypes.SECURETYPE,
    tbl1Employees.EmployeeType_ID,
    tbl3EmployeeAccess.HaveAccess,
    tbl3EmployeeAccess.FormName

FROM
    (tbl5EmployeeSecurityTypes INNER JOIN (tbl2EmployeeTypes INNER JOIN tbl1Employees ON tbl2EmployeeTypes.ID = tbl1Employees.EmployeeType_ID) ON tbl5EmployeeSecurityTypes.SECUREID = tbl2EmployeeTypes.SECUREID)
    INNER JOIN tbl3EmployeeAccess
        ON tbl5EmployeeSecurityTypes.SECUREID = tbl3EmployeeAccess.SECUREID;
