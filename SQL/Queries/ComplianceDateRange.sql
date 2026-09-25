SELECT
    DATETABLE.Date

FROM
    DATETABLE

WHERE
    (((DATETABLE.Date) Between #1/1/2021# And DateSerial(Year(Date()),Month(Date()),0)));
