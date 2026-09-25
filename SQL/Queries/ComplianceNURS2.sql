SELECT
First(DATETABLE.Month) AS FirstOfMonth,
    DATETABLE.Year,
    ComplianceNURS1.QMDATE

FROM
    ComplianceNURS1
    LEFT JOIN DATETABLE
        ON ComplianceNURS1.QMDATE = DATETABLE.Date

GROUP BY
    DATETABLE.Year
    ComplianceNURS1.QMDATE;
