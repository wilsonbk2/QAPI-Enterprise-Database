SELECT
Max(tblUserLog.Timestamp) AS MaxOfTimestamp,
    tblUserLog.UserComputerName

FROM
    tblUserLog

GROUP BY
    tblUserLog.UserComputerName

HAVING
    (((tblUserLog.UserComputerName) Is Not Null));
