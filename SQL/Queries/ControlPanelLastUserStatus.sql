SELECT
    [ControlPanelCurrentActiveUsers 1].MaxOfTimestamp,
    [ControlPanelCurrentActiveUsers 1].UserComputerName,
    tbl1Employees.FirstName,
    tbl1Employees.LastName,
    tbl1Employees.Alias,
    tblUserLog.ComputerName,
    tblUserLog.DomainName,
    tblUserLog.Accessed,
    tblUserLog.Activity

FROM
    tbl1Employees
    INNER JOIN ([ControlPanelCurrentActiveUsers 1] INNER JOIN tblUserLog
        ON ([ControlPanelCurrentActiveUsers 1].UserComputerName = tblUserLog.UserComputerName)
        AND
        ([ControlPanelCurrentActiveUsers 1].MaxOfTimestamp = tblUserLog.Timestamp)) ON tbl1Employees.Alias = tblUserLog.UserComputerName

ORDER BY
    [ControlPanelCurrentActiveUsers 1].MaxOfTimestamp DESC;
