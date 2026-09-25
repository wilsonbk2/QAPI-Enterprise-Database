SELECT
    TOP 1 StayInterviewsRandomizedList.[Full Name],
Rnd([Rand]*Timer()*-1) AS Y,
    StayInterviewsRandomizedList.[EMP ID],
    EmployeeOnlyCurrent.JOBTITLE,
    EmployeeOnlyCurrent.[CC-BASE],
    EmployeeOnlyCurrent.[BLDG#],
    EmployeeOnlyCurrent.POD,
    EmployeeOnlyCurrent.WARD,
    EmployeeOnlyCurrent.[Supervisor Name]

FROM
    StayInterviewsRandomizedList
    LEFT JOIN EmployeeOnlyCurrent
        ON StayInterviewsRandomizedList.[EMP ID] = EmployeeOnlyCurrent.[EMP ID]

ORDER BY
    Rnd([Rand]*Timer()*-1);
