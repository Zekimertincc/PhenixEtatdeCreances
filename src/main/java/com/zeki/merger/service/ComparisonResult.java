package com.zeki.merger.service;

import java.util.List;

public record ComparisonResult(
    List<DiffRow>           allRows,
    List<DiffRow>           discrepancies,
    List<UnmatchedProcRow>  unmatchedProcreances,
    List<UnmatchedConsoRow> unmatchedConso
) {}
