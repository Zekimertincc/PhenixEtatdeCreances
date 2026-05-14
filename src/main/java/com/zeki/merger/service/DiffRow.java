package com.zeki.merger.service;

public record DiffRow(
    String  clientName,
    String  clientCode,
    double  procHono,
    double  consoCommTtc,
    double  diff,
    boolean hasDiscrepancy
) {}
