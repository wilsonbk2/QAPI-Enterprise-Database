SELECT
    ADMIS.DEPTID,
    ADMIS.TOPICID,
    Format(
        [QMDEC],
        "Percent"
    )AS [PERCENT],
    ADMIS.QMDATE

FROM
    ADMIS

ORDER BY
    ADMIS.QMDATE;
