SELECT
    StaffingLogReport.SDATE,
    StaffingLogReport.DESC,
    StaffingLogReport.CSCDESC,
    StaffingLogReport.UNITMAIN,
    StaffingLogReport.UNITSUB,
    StaffingLogReport.JOBTITLE

FROM
    StaffingLogReport

GROUP BY
    StaffingLogReport.SDATE
    StaffingLogReport.DESC
    StaffingLogReport.CSCDESC
    StaffingLogReport.UNITMAIN
    StaffingLogReport.UNITSUB
    StaffingLogReport.JOBTITLE

ORDER BY
    StaffingLogReport.DESC DESC
    StaffingLogReport.CSCDESC
    StaffingLogReport.UNITMAIN
    StaffingLogReport.UNITSUB
    StaffingLogReport.JOBTITLE PIVOT StaffingLogReport.SHIFT In ("DAY","EVENING","NIGHT");
