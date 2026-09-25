SELECT
    ComplianceCurrentDateRangeMonths.Year,
    ComplianceCurrentDateRangeMonths.Month,
Val(Replace([MonthNo],"-","")) AS MonNum,
    IIf(
IsNull([FirstOfMonth]),
        "No Data Submission",
        "Successful Submission"
    )AS [Submit Status]

FROM
    ComplianceCurrentDateRangeMonths
    LEFT JOIN ComplianceHIM
        ON (ComplianceCurrentDateRangeMonths.Month = ComplianceHIM.FirstOfMonth)
        AND
        (ComplianceCurrentDateRangeMonths.Year = ComplianceHIM.Year)

ORDER BY
    ComplianceCurrentDateRangeMonths.Year
    Val(Replace([MonthNo],"-",""));
