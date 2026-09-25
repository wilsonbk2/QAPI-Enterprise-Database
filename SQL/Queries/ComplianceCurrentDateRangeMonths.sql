SELECT
    DATETABLE.Year,
    DATETABLE.Month,
    DATETABLE.MonthNo

FROM
    ComplianceDateRange
    INNER JOIN DATETABLE
        ON ComplianceDateRange.Date = DATETABLE.Date

GROUP BY
    DATETABLE.Year
    DATETABLE.Month
    DATETABLE.MonthNo

ORDER BY
    DATETABLE.Year DESC
    DATETABLE.MonthNo DESC;
