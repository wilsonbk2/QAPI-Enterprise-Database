SELECT
    [LNAME] & ", " & [FNAME] & " " & [MI] & " " & [patepi] AS FullName,
    DEQPATIENT.patepi

FROM
    DEQPATIENT

ORDER BY
    [LNAME] & ", " & [FNAME] & " " & [MI] & " " & [patepi];
