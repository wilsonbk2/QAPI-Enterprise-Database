SELECT
    tbl1Employees.EMPID,
    SATELLITE2.[Full Name],
    tbl4UserAlias.UserName,
    [ControlPanelCurrentActiveUsers 4].MaxOfTimestamp

FROM
    [ControlPanelCurrentActiveUsers 4]
    RIGHT JOIN ((SATELLITE2 RIGHT JOIN tbl1Employees
        ON SATELLITE2.[EMP ID] = tbl1Employees.EMPID) RIGHT JOIN tbl4UserAlias ON tbl1Employees.Alias = tbl4UserAlias.Alias) ON ([ControlPanelCurrentActiveUsers 4].ComputerName = tbl4UserAlias.UserName)
        AND
        ([ControlPanelCurrentActiveUsers 4].UserComputerName = tbl4UserAlias.Alias)

WHERE
    ((([ControlPanelCurrentActiveUsers 4].MaxOfTimestamp) Is Not Null))

ORDER BY
    SATELLITE2.[Full Name]
    [ControlPanelCurrentActiveUsers 4].MaxOfTimestamp DESC;
