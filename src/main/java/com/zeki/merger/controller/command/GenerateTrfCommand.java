package com.zeki.merger.controller.command;

import com.zeki.merger.core.exception.BusinessException;
import com.zeki.merger.core.exception.ErrorCode;
import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.service.util.ProgressNotifier;
import com.zeki.merger.trf.TrfGeneratorService;

import java.io.File;
import java.util.Map;

/**
 * Command: generates the TRF workbook.
 *
 * Required context keys: consoFile, listingFile, tableauFile, outputFolder.
 */
public class GenerateTrfCommand implements ReportCommand {

    private final TrfGeneratorService trfGeneratorService;

    public GenerateTrfCommand(DatabaseManager dbManager) {
        this.trfGeneratorService = new TrfGeneratorService(dbManager);
    }

    @Override
    public File execute(Map<String, Object> context, ProgressNotifier notifier) throws Exception {
        File consoFile    = requireFile(context, "consoFile");
        File listingFile  = requireFile(context, "listingFile");
        File tableauFile  = requireFile(context, "tableauFile");
        File outputFolder = requireFolder(context, "outputFolder");

        return trfGeneratorService.generate(
            consoFile, listingFile, tableauFile, outputFolder,
            notifier.asBiConsumer()
        );
    }

    @Override
    public String getName() {
        return "GENERATE_TRF";
    }

    private File requireFile(Map<String, Object> ctx, String key) {
        Object val = ctx.get(key);
        if (!(val instanceof File f) || !f.exists()) {
            throw new BusinessException(
                ErrorCode.FILE_NOT_FOUND,
                "Required context key missing or file not found: " + key
            );
        }
        return f;
    }

    private File requireFolder(Map<String, Object> ctx, String key) {
        Object val = ctx.get(key);
        if (!(val instanceof File f) || !f.isDirectory()) {
            throw new BusinessException(
                ErrorCode.FILE_NOT_FOUND,
                "Required output folder missing: " + key
            );
        }
        return f;
    }
}
