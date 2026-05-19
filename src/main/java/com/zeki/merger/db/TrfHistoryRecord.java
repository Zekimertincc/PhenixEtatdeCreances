package com.zeki.merger.db;

public record TrfHistoryRecord(
        long    id,
        long    monthId,
        String  clientName,
        String  clientCode,
        double  encaissements,
        double  montantFacturer,
        double  nousDoit,
        double  sommesReverser,
        String  etat,
        String  iban,
        boolean nonCompensation) {}
