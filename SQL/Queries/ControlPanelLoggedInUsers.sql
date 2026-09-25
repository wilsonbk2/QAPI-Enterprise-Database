SELECT
    [ControlPanelCurrentActiveUsers 3].MaxOfTimestamp,
    [ControlPanelCurrentActiveUsers 3].UserComputerName,
    tbl1Employees.FirstName,
    tbl1Employees.LastName,
    tbl1Employees.Alias,
    tblUserLog.ComputerName,
    tblUserLog.DomainName,
    tblUserLog.Accessed,
    tblUserLog.Activity

FROM
    [ControlPanelCurrentActiveUsers 3]
    INNER JOIN (tbl1Employees INNER JOIN tblUserLog
        ON tbl1Employees.Alias = tblUserLog.UserComputerName) ON ([ControlPanelCurrentActiveUsers 3].UserComputerName = tblUserLog.UserComputerName)
        AND
        ([ControlPanelCurrentActiveUsers 3].MaxOfTimestamp = tblUserLog.Timestamp)

WHERE
    (((tblUserLog.Activity)="Login QAPI"))

ORDER BY
    [ControlPanelCurrentActiveUsers 3].MaxOfTimestamp DESC;
