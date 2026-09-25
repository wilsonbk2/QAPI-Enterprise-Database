SELECT
Max([ControlPanelCurrentActiveUsers 2].Timestamp) AS MaxOfTimestamp,
    [ControlPanelCurrentActiveUsers 2].UserComputerName

FROM
    [ControlPanelCurrentActiveUsers 2]

GROUP BY
    [ControlPanelCurrentActiveUsers 2].UserComputerName

HAVING
    ((([ControlPanelCurrentActiveUsers 2].UserComputerName) Is Not Null));
