package com.zeki.merger.controller.command;

import com.zeki.merger.core.exception.BusinessException;
import com.zeki.merger.core.exception.ErrorCode;
import com.zeki.merger.db.DatabaseManager;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * Factory that maps command names to ReportCommand instances.
 * New commands can be registered at runtime via register().
 */
public class ReportCommandFactory {

    private final Map<String, ReportCommand> commands = new HashMap<>();

    public ReportCommandFactory(DatabaseManager dbManager) {
        register(new GenerateTrfCommand(dbManager));
        // Additional commands wired here as they are extracted from MainController:
        // register(new GenerateEtatPublicCommand());
        // register(new CompareFilesCommand());
        // register(new FixPathsCommand());
        // register(new RunConsolidationCommand(dbManager));
    }

    public void register(ReportCommand command) {
        commands.put(command.getName().toUpperCase(), command);
    }

    public ReportCommand getCommand(String name) {
        ReportCommand cmd = commands.get(name.toUpperCase());
        if (cmd == null) {
            throw new BusinessException(
                ErrorCode.UNKNOWN_COMMAND,
                "Unknown command: " + name,
                Map.of("requested", name, "available", getCommandNames())
            );
        }
        return cmd;
    }

    public Set<String> getCommandNames() {
        return commands.keySet();
    }
}
