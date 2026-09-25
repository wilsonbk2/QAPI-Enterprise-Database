SELECT
    BUGS.Alias,
    BUGS.BTime,
    BUGS.BReason,
    BUGS.BComment,
    BUGS.Resolved

FROM
    BUGS

WHERE
    (((BUGS.Resolved)="Needs Attention"));
