SELECT
    SATELLITE2.[Full Name],
    tbl6EmployeeMultipleTypes.EMPID,
    tbl6EmployeeMultipleTypes.Alias,
    tbl2EmployeeTypes.EmployeeType,
    tbl2EmployeeTypes.IDD

FROM
    SATELLITE2
    RIGHT JOIN (tbl2EmployeeTypes RIGHT JOIN tbl6EmployeeMultipleTypes
        ON tbl2EmployeeTypes.IDD = tbl6EmployeeMultipleTypes.IDD) ON SATELLITE2.[EMP ID] = tbl6EmployeeMultipleTypes.EMPID;
