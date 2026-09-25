SELECT
    DEQPATIENT.PatientID,
    DEQPATIENT.EPI,
    [LNAME] & ", " & [FNAME] & " " & [MI] & "." AS [Full Name],
    BUILDING.[BUILDING#],
    UNIT.UNITMAIN,
    DEQPATIENT.UNIT,
    DEQPATIENT.ADMISDT,
    DEQPATIENT.FNAME,
    DEQPATIENT.LNAME,
    DEQPATIENT.MI,
Val([Deliquencies]) AS Delinquencies,
    IIf(
IsNull([DISDT]),
        "Current",
        [DISDT]
    )AS DisStat,
    DEQPATIENT.DISDT

FROM
    ((DeliquentPatientMostCurrent INNER JOIN (UNIT RIGHT JOIN DEQPATIENT ON UNIT.UNITSUB = DEQPATIENT.UNIT) ON (DeliquentPatientMostCurrent.MaxOfEPI = DEQPATIENT.EPI) AND (DeliquentPatientMostCurrent.PatientID = DEQPATIENT.PatientID)) LEFT JOIN BUILDING ON UNIT.BUI = BUILDING.BUI)
    LEFT JOIN [DeliquencyDefCount 2]
        ON DEQPATIENT.PatientID = [DeliquencyDefCount 2].PatientID

ORDER BY
    Val([Deliquencies]) DESC;
