SELECT
    SATELLITE2.[EMP ID],
    SATELLITE2.[Full Name],
    SATELLITE2.BIRTHDATE,
    SATELLITE2.[Photo Path],
    SATELLITE2.JOBTITLE,
Day([BIRTHDATE]) AS [Day],
    SATELLITE2.EMPSTATUS,
Month([BIRTHDATE]) AS [Month],
    Format(
        [BIRTHDATE],
        "mmm dd"
    )AS BIRTHDATE2

FROM
    SATELLITE2

WHERE
    (((SATELLITE2.EMPSTATUS)="Current") AND ((Month([BIRTHDATE]))=Month(Date())))

ORDER BY
    Day([BIRTHDATE]);
