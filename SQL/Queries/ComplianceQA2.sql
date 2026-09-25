SELECT
    DATETABLE.Month,
    DATETABLE.Year

FROM
    ComplianceQA1
    INNER JOIN DATETABLE
        ON ComplianceQA1.QMDATE = DATETABLE.Date;
