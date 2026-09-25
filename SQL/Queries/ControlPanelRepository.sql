SELECT
    tbl1Employees.EMPID,
    SATELLITE2.[EMP ID],
    SATELLITE2.FIRSTNAME,
    SATELLITE2.LASTNAME,
    SATELLITE2.MI,
    tbl1Employees.Alias,
    tbl1Employees.EmployeeType_ID,
    tbl5EmployeeSecurityTypes.SECURETYPE,
    SATELLITE2.EMPSTATUS,
    SATELLITE2.[Photo Path]

FROM
    (tbl5EmployeeSecurityTypes INNER JOIN (tbl2EmployeeTypes INNER JOIN tbl1Employees ON tbl2EmployeeTypes.ID = tbl1Employees.EmployeeType_ID) ON tbl5EmployeeSecurityTypes.SECUREID = tbl2EmployeeTypes.SECUREID)
    LEFT JOIN SATELLITE2
        ON tbl1Employees.EMPID = SATELLITE2.[EMP ID];
