SELECT
    UNIT.UNITSUB,
    UNIT.UNITMAIN

FROM
    UNIT

WHERE
    (((UNIT.BUI)=[Forms]![CODElogging]![B]))

ORDER BY
    UNIT.UNITMAIN;
