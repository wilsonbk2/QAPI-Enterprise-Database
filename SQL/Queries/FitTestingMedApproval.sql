SELECT
    N95APPRV.EMPID,
    N95APPRV.N95APPRV AS APP,
    N95APPRV.MEDSCREEN,
    N95APPRV.N95DATE,
    IIf(
        [APP]="NO",
        "Denied for N95",
        "Screened and Fitted"
    )AS APPROVAL

FROM
    FitTestingMostCurrentMedScreen
    INNER JOIN N95APPRV
        ON (FitTestingMostCurrentMedScreen.MaxOfN95DATE = N95APPRV.N95DATE)
        AND
        (FitTestingMostCurrentMedScreen.EMPID = N95APPRV.EMPID);
