SELECT
    tbl7forensicdatalog.DEQKEY,
    tbl7forensicdatalog.AREA,
    tbl7forensicdatalog.BGVALUE,
    tbl7forensicdatalog.NEWVALUE,
    SATELLITE2.[Full Name],
    tbl7forensicdatalog.Timestamp

FROM
    (tbl7forensicdatalog LEFT JOIN tbl1Employees ON tbl7forensicdatalog.Alias = tbl1Employees.Alias)
    LEFT JOIN SATELLITE2
        ON tbl1Employees.EMPID = SATELLITE2.[EMP ID]

WHERE
    (((tbl7forensicdatalog.AREA)="Edit Class"))

ORDER BY
    tbl7forensicdatalog.Timestamp DESC;
