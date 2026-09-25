SELECT
    epicurr.patidnum,
Max(epicurr.episode) AS MaxOfepisode,
Max(epicurr.dischrgdt) AS MaxOfdischrgdt,
    patinfo.firstname,
    patinfo.lastname

FROM
    epicurr
    INNER JOIN patinfo
        ON epicurr.patidnum = patinfo.patidnum

GROUP BY
    epicurr.patidnum
    patinfo.firstname
    patinfo.lastname;
