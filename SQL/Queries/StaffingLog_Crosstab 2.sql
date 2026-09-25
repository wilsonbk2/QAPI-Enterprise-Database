SELECT
    StaffingLogReportRange.SDATE,
    StaffingLogReportRange.DESC,
    StaffingLogReportRange.CSCDESC,
    StaffingLogReportRange.UNITMAIN,
    StaffingLogReportRange.UNITSUB,
    StaffingLogReportRange.JOBTITLE

FROM
    StaffingLogReportRange

GROUP BY
    StaffingLogReportRange.SDATE
    StaffingLogReportRange.DESC
    StaffingLogReportRange.CSCDESC
    StaffingLogReportRange.UNITMAIN
    StaffingLogReportRange.UNITSUB
    StaffingLogReportRange.JOBTITLE

ORDER BY
    StaffingLogReportRange.DESC DESC
    StaffingLogReportRange.CSCDESC
    StaffingLogReportRange.UNITMAIN
    StaffingLogReportRange.UNITSUB
    StaffingLogReportRange.JOBTITLE PIVOT StaffingLogReportRange.SHIFT In ("DAY","EVENING","NIGHT");
