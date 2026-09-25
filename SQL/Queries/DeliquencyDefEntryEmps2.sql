SELECT
    DeliquencyDefEntryEmps.Alias,
    SATELLITE2.[Full Name]

FROM
    DeliquencyDefEntryEmps
    INNER JOIN (tbl1Employees INNER JOIN SATELLITE2
        ON tbl1Employees.EMPID = SATELLITE2.[EMP ID]) ON DeliquencyDefEntryEmps.Alias = tbl1Employees.Alias;
