SELECT
    SATELLITE2.[EMP ID],
    tbl1Employees.EMPID,
    tbl1Employees.Alias,
    SATELLITE2.FIRSTNAME,
    SATELLITE2.LASTNAME,
    SATELLITE2.[MI] AS MiddleInitial,
    SATELLITE2.[Full Name]

FROM
    tbl1Employees
    LEFT JOIN SATELLITE2
        ON tbl1Employees.EMPID = SATELLITE2.[EMP ID]

WHERE
    (((SATELLITE2.FIRSTNAME) Is Not Null));
