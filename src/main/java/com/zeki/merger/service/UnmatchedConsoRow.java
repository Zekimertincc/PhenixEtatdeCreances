package com.zeki.merger.service;

public record UnmatchedConsoRow(
    String name,
    String code,
    double commissions,
    double sommesCz,
    double sommesReverser
) {}
