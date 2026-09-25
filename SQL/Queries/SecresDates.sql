SELECT
    SECRES.[Patient ID],
    SECRES.EventDate

FROM
    SECRES

WHERE
    (((SECRES.EventDate) Between DateSerial(Year(Date()),Month(Date())-1,1) And DateSerial(Year(Date()),Month(Date()),0)));
