SELECT
    UNIZIP.state_id,
    UNIZIP.state_name

FROM
    UNIZIP

GROUP BY
    UNIZIP.state_id
    UNIZIP.state_name;
