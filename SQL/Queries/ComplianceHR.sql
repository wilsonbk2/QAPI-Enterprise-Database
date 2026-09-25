SELECT
First(DATETABLE.Month) AS FirstOfMonth,
    DATETABLE.Year,
    HREmployee.HRDATE

FROM
    HREmployee
    INNER JOIN DATETABLE
        ON HREmployee.HRDATE = DATETABLE.Date

GROUP BY
    DATETABLE.Year
    HREmployee.HRDATE;
