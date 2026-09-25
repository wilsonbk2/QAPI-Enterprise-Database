SELECT
    tbl2EmployeeTypes.ID,
    tbl2EmployeeTypes.EmployeeType

FROM
    tbl2EmployeeTypes

WHERE
    (((tbl2EmployeeTypes.ID)<>3 And (tbl2EmployeeTypes.ID)<>23))

ORDER BY
    tbl2EmployeeTypes.EmployeeType;
