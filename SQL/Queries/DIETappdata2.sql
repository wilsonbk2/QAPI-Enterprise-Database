SELECT
    DIET.DEPTID,
    DIET.TOPICID,
    DIETappdata1.QMVALUE AS QMTOTAL,
    [DIET].[QMVALUE]/[QMTOTAL] AS [DECIMAL],
    Format(
        [DECIMAL],
        "Percent"
    )AS [PERCENT],
    DIET.QMDATE

FROM
    DIETappdata1
    INNER JOIN DIET
        ON DIETappdata1.QMDATE = DIET.QMDATE

WHERE
    (((DIET.TOPICID)<>"00048"));
