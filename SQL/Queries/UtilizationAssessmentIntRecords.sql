SELECT
    UtilizationInitialPat.PID,
    UtilizationInitialPat.Episode,
    DEQPATIENT.PatientID,
    DEQPATIENT.EPI,
    [LNAME] & ", " & [FNAME] & " " & [MI] AS [Patient Name],
    DEQPATIENT.ADMISDT,
    DEQPATIENT.ADMISTM,
    IIf(
IsNull([PatientID]),
        "47",
        "47"
    )AS [Psychosocial Code],
    IIf(
IsNull([PatientID]),
        "44",
        "44"
    )AS [Px Rec Code],
    IIf(
IsNull([PatientID]),
        "109",
        "109"
    )AS [Rehab CorRehab Code],
    IIf(
IsNull([PatientID]),
        "110",
        "110"
    )AS Nutrition,
    IIf(
IsNull([PatientID]),
        "97",
        "97"
    )AS Dysphagia,
    IIf(
IsNull([PatientID]),
        "111",
        "111"
    )AS [Falls Code],
    IIf(
IsNull([PatientID]),
        "43",
        "43"
    )AS [H/P Code],
    IIf(
IsNull([PatientID]),
        "49",
        "49"
    )AS [CSSRS - Nursing Code],
    IIf(
IsNull([PatientID]),
        "113",
        "113"
    )AS [CSSRS - LIP Code],
    IIf(
IsNull([PatientID]),
        "72",
        "72"
    )AS [Pain - Nursing Code],
    IIf(
IsNull([PatientID]),
        "112",
        "112"
    )AS [Pain - Tx Plan Code],
    IIf(
IsNull([PatientID]),
        "45",
        "45"
    )AS [BH Intake Code],
    IIf(
IsNull([PatientID]),
        "104",
        "104"
    )AS [Nursing Triage Code],
    IIf(
IsNull([PatientID]),
        "102",
        "102"
    )AS [Psychology Code],
    IIf(
IsNull([PatientID]),
        "46",
        "46"
    )AS [Psychiatric Code],
    DEQPATIENT.patepi,
    UtilizationInitialPat.BUI,
    UtilizationInitialPat.UNIT

FROM
    DEQPATIENT
    LEFT JOIN UtilizationInitialPat
        ON (DEQPATIENT.EPI = UtilizationInitialPat.Episode)
        AND
        (DEQPATIENT.PatientID = UtilizationInitialPat.PID)

WHERE
    (((UtilizationInitialPat.PID) Is Not Null) AND ((DEQPATIENT.ADMISDT)>#6/1/2021#));
