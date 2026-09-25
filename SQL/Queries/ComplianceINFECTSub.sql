SELECT
    ComplianceCurrentDateRangeMonths.Year,
    ComplianceCurrentDateRangeMonths.Month,
Val(Replace([MonthNo],"-","")) AS MonNum,
    IIf(
IsNull([ComplianceINFECT].[Month]),
        "No Data Submission",
        "Successful Submission"
    )AS [Submit Status]

FROM
    ComplianceCurrentDateRangeMonths
    LEFT JOIN ComplianceINFECT
        ON (ComplianceCurrentDateRangeMonths.Year = ComplianceINFECT.Year)
        AND
        (ComplianceCurrentDateRangeMonths.Month = ComplianceINFECT.Month)

ORDER BY
    Val(Replace([MonthNo],"-",""));
