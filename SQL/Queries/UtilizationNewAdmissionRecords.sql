SELECT
    DEQPATIENT.ADMISDT,
    DEQPATIENT.patepi,
    [LNAME] & ", " & [FNAME] & " " & [MI] AS Fullname,
    DEQPATIENT.BUI,
    DEQPATIENT.UNIT,
    DEQPATIENT.ADMISTM,
    UtilizationAssessmentMainStatus.PatientID,
    DEQPATIENT.PatientID,
    DEQPATIENT.EPI,
    DEQPATIENT.FNAME,
    DEQPATIENT.LNAME,
    DEQPATIENT.MI

FROM
    DEQPATIENT
    LEFT JOIN UtilizationAssessmentMainStatus
        ON (DEQPATIENT.EPI = UtilizationAssessmentMainStatus.EPI)
        AND
        (DEQPATIENT.PatientID = UtilizationAssessmentMainStatus.PatientID)

WHERE
    (((DEQPATIENT.ADMISDT) Between #7/1/2021# And Now()) AND ((UtilizationAssessmentMainStatus.PatientID) Is Null))

ORDER BY
    DEQPATIENT.ADMISDT DESC
    DEQPATIENT.ADMISTM DESC;
