package com.zeki.merger.db;

public record TrfMonthRecord(
        long   id,
        int    year,
        int    month,
        String status,
        int    nbClients,
        double totalMontant,
        double totalNousDoit,
        String closedAt) {}
