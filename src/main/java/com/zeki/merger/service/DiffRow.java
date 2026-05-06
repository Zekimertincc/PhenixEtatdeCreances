package com.zeki.merger.service;

public record DiffRow(
    String  clientName,
    String  clientCode,
    double  procHonoTtc,
    double  consoCommissions,
    double  diffHono,
    double  procDisponible,
    double  consoSommesCz,
    double  diffDisponible,
    double  procReversement,
    double  consoSommesReverser,
    double  diffReversement,
    boolean hasDiscrepancy
) {}
