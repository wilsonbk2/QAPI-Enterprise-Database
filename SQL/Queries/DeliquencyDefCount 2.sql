SELECT
    PatientMasterIDList.PatientID,
    DeliquencyDefCount.CountOfPatientID,
    IIf(
IsNull([CountOfPatientID]),
        "0",
        [CountOfPatientID]
    )AS Deliquencies

FROM
    PatientMasterIDList
    LEFT JOIN DeliquencyDefCount
        ON PatientMasterIDList.PatientID = DeliquencyDefCount.PatientID

ORDER BY
    DeliquencyDefCount.CountOfPatientID DESC;
