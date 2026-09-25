SELECT
Max(tblUserLog.Timestamp) AS MaxOfTimestamp,
    tblUserLog.UserComputerName,
    tblUserLog.ComputerName

FROM
    tblUserLog

GROUP BY
    tblUserLog.UserComputerName
    tblUserLog.ComputerName

HAVING
    (((tblUserLog.UserComputerName) Is Not Null));
