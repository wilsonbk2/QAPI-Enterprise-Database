SELECT
    TrainingViewIgnoredClassesCrosstab.[Content Title],
    TrainingViewIgnoredClassesCrosstab.[Content Type],
    TrainingViewIgnoredClassesCrosstab.COURSEID,
    TrainingViewIgnoredClassesCrosstab.[0000]

FROM
    TrainingViewIgnoredClassesCrosstab

ORDER BY
    TrainingViewIgnoredClassesCrosstab.[0000] DESC;
