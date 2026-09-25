SELECT
    EmployeeOnlyALL.[EMP ID] AS EID,
    EmployeeOnlyALL.[Full Name],
    EmployeeOnlyALL.EMPSTATUS,
    IIf(
IsNull([COVIDDATE]),
        "No Record",
        IIf(
Date()>[COVIDDATE]+365,
            "Expired",
            "Current"
        )
    )AS [COVID-19],
    INFECTvacc1.VACDESC AS COVID,
    INFECTvacc1.VACDATE AS COVIDDATE,
    IIf(
IsNull([HEPADATE]),
        "No Record",
        "Current"
    )AS [Hepatitis A],
    INFECTvacc2.VACDESC AS HEPA,
    INFECTvacc2.VACDATE AS HEPADATE,
    IIf(
IsNull([HEPBDATE]),
        "No Record",
        "Current"
    )AS [Hepatitis B],
    INFECTvacc3.VACDESC AS HEPB,
    INFECTvacc3.VACDATE AS HEPBDATE,
    IIf(
IsNull([FLUDATE]),
        "No Record",
        IIf(
Date()>[FLUDATE]+365,
            "Expired",
            "Current"
        )
    )AS Influenza,
    INFECTvacc4.VACDESC AS FLU,
    INFECTvacc4.VACDATE AS FLUDATE,
    IIf(
IsNull([MMRDATE]),
        "No Record",
        "Current"
    )AS MMR,
    INFECTvacc5.VACDESC AS MMRV,
    INFECTvacc5.VACDATE AS MMRDATE,
    IIf(
IsNull([PREVDATE]),
        "No Record",
        "Current"
    )AS Prevnar,
    INFECTvacc6.VACDESC AS PREV,
    INFECTvacc6.VACDATE AS PREVDATE,
    IIf(
IsNull([PNEDATE]),
        "No Record",
        "Current"
    )AS Pneumovax,
    INFECTvacc7.VACDESC AS PNE,
    INFECTvacc7.VACDATE AS PNEDATE,
    IIf(
IsNull([SHINGDATE]),
        "No Record",
        "Current"
    )AS Shingles,
    INFECTvacc9.VACDESC AS SHING,
    INFECTvacc9.VACDATE AS SHINGDATE,
    IIf(
IsNull([TDAPDATE]),
        "No Record",
        IIf(
Date()>[TDAPDATE]+3650,
            "Expired",
            "Current"
        )
    )AS Tdap,
    INFECTvacc010.VACDESC AS TDAPV,
    INFECTvacc010.VACDATE AS TDAPDATE,
    IIf(
IsNull([TSTDATE]),
        "No Record",
        IIf(
Date()>[TSTDATE]+365,
            "Expired",
            "Current"
        )
    )AS Tuberculosis,
    INFECTvacc011.VACDESC AS TST,
    INFECTvacc011.VACDATE AS TSTDATE,
    IIf(
IsNull([VARIDATE]),
        "No Record",
        "Current"
    )AS Varicella,
    INFECTvacc012.VACDESC AS VARI,
    INFECTvacc012.VACDATE AS VARIDATE,
    "1" AS [COVID-19KEY],
    "2" AS HepatitisAKEY,
    "3" AS HepatitisBKEY,
    "4" AS InfluenzaKEY,
    "5" AS MMRKEY,
    "6" AS PrevnarKEY,
    "7" AS PneumovaxKEY,
    "9" AS ShinglesKEY,
    "10" AS TdapKEY,
    "11" AS TuberculosisKEY,
    "12" AS VaricellaKEY,
    EmployeeOnlyALL.POD,
    EmployeeOnlyALL.WARD,
    EmployeeOnlyALL.[BLDG#],
    BUILDING.[BUILDING#]

FROM
    (((((((((((EmployeeOnlyALL LEFT JOIN INFECTvacc1 ON EmployeeOnlyALL.[EMP ID] = INFECTvacc1.EID) LEFT JOIN INFECTvacc2 ON EmployeeOnlyALL.[EMP ID] = INFECTvacc2.EID) LEFT JOIN INFECTvacc3 ON EmployeeOnlyALL.[EMP ID] = INFECTvacc3.EID) LEFT JOIN INFECTvacc4 ON EmployeeOnlyALL.[EMP ID] = INFECTvacc4.EID) LEFT JOIN INFECTvacc5 ON EmployeeOnlyALL.[EMP ID] = INFECTvacc5.EID) LEFT JOIN INFECTvacc6 ON EmployeeOnlyALL.[EMP ID] = INFECTvacc6.EID) LEFT JOIN INFECTvacc7 ON EmployeeOnlyALL.[EMP ID] = INFECTvacc7.EID) LEFT JOIN INFECTvacc9 ON EmployeeOnlyALL.[EMP ID] = INFECTvacc9.EID) LEFT JOIN INFECTvacc010 ON EmployeeOnlyALL.[EMP ID] = INFECTvacc010.EID) LEFT JOIN INFECTvacc011 ON EmployeeOnlyALL.[EMP ID] = INFECTvacc011.EID) LEFT JOIN INFECTvacc012 ON EmployeeOnlyALL.[EMP ID] = INFECTvacc012.EID)
    INNER JOIN BUILDING
        ON EmployeeOnlyALL.[BLDG#] = BUILDING.BUI

WHERE
    (((EmployeeOnlyALL.[EMP ID])<>"000000000") AND ((EmployeeOnlyALL.[Full Name]) Is Not Null))

ORDER BY
    EmployeeOnlyALL.[Full Name];
