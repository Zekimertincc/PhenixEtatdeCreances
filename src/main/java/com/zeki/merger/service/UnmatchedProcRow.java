package com.zeki.merger.service;

public record UnmatchedProcRow(
    String name,
    String code,
    double honoTtc,
    double disponible,
    double reversement
) {}
