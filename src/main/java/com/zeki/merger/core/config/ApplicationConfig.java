package com.zeki.merger.core.config;

import com.zeki.merger.controller.command.ReportCommandFactory;
import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.service.io.DataIOFactory;
import com.zeki.merger.service.report.ReportStrategyFactory;

/**
 * Singleton configuration hub.
 *
 * Centralizes object creation so services don't instantiate their own dependencies.
 * Replaces the scattered `new XyzService()` calls that live directly in MainController.
 *
 * Usage:
 *   ApplicationConfig cfg = ApplicationConfig.getInstance();
 *   cfg.getReportCommandFactory().getCommand("GENERATE_TRF").execute(ctx, notifier);
 */
public class ApplicationConfig {

    private static volatile ApplicationConfig instance;

    private final DatabaseManager        databaseManager;
    private final ReportStrategyFactory  reportStrategyFactory;
    private final DataIOFactory          dataIOFactory;
    private final ReportCommandFactory   reportCommandFactory;

    private ApplicationConfig() {
        this.databaseManager       = DatabaseManager.getInstance();
        this.reportStrategyFactory = new ReportStrategyFactory();
        this.dataIOFactory         = new DataIOFactory();
        this.reportCommandFactory  = new ReportCommandFactory(databaseManager);
    }

    public static ApplicationConfig getInstance() {
        if (instance == null) {
            synchronized (ApplicationConfig.class) {
                if (instance == null) {
                    instance = new ApplicationConfig();
                }
            }
        }
        return instance;
    }

    public DatabaseManager       getDatabaseManager()       { return databaseManager;       }
    public ReportStrategyFactory getReportStrategyFactory() { return reportStrategyFactory; }
    public DataIOFactory         getDataIOFactory()         { return dataIOFactory;          }
    public ReportCommandFactory  getReportCommandFactory()  { return reportCommandFactory;  }
}
