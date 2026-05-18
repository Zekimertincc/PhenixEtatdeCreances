package com.zeki.merger.service;

import java.io.File;
import java.io.IOException;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.*;
import java.util.concurrent.TimeUnit;
import java.util.function.BiConsumer;

import static java.nio.file.StandardWatchEventKinds.*;

public class FolderWatchService {

    private final EtatCreancesSyncService    syncService;
    private final FolderScanner              scanner;
    private final BiConsumer<String, String> onEvent;

    private Thread           watchThread;
    private volatile boolean running = false;
    private File             rootFolder;
    private final Map<Path, Long> pending = new HashMap<>();
    private final Map<Path, Long> lastModified = new HashMap<>();
    private long lastPollTime = 0;

    public FolderWatchService(EtatCreancesSyncService syncService,
                               BiConsumer<String, String> onEvent) {
        this.syncService = syncService;
        this.scanner     = new FolderScanner();
        this.onEvent     = onEvent;
    }

    public synchronized void start(File root) {
        if (running) return;
        this.rootFolder = root;
        pending.clear();
        lastModified.clear();
        lastPollTime = 0;
        running = true;
        watchThread = Thread.ofVirtual().name("folder-watcher").start(() -> {
            try (WatchService ws = FileSystems.getDefault().newWatchService()) {
                registerAll(root.toPath(), ws);

                while (running) {
                    WatchKey key = ws.poll(500, TimeUnit.MILLISECONDS);
                    if (key == null) {
                        flushPending();
                        if (System.currentTimeMillis() - lastPollTime > 30_000) {
                            lastPollTime = System.currentTimeMillis();
                            pollCompanyFiles();
                        }
                        continue;
                    }
                    for (WatchEvent<?> event : key.pollEvents()) {
                        if (event.kind() == OVERFLOW) continue;
                        Path changed = ((Path) key.watchable())
                                           .resolve((Path) event.context());
                        String name = changed.getFileName().toString().toLowerCase();

                        if ((name.endsWith(".xlsx") || name.endsWith(".xls"))
                                && !name.startsWith("~$")) {
                            pending.put(changed, System.currentTimeMillis() + 5000);
                        }
                        if (event.kind() == ENTRY_CREATE && changed.toFile().isDirectory()) {
                            registerAll(changed, ws);
                        }
                    }
                    key.reset();
                    flushPending();
                    if (System.currentTimeMillis() - lastPollTime > 30_000) {
                        lastPollTime = System.currentTimeMillis();
                        pollCompanyFiles();
                    }
                }
            } catch (Exception e) {
                if (running) onEvent.accept("SYSTEM", "Erreur surveillance : " + e.getMessage());
            }
        });
    }

    public synchronized void stop() {
        running = false;
        if (watchThread != null) {
            watchThread.interrupt();
            watchThread = null;
        }
    }

    public boolean isRunning() { return running; }

    // -------------------------------------------------------------------------

    private void flushPending() {
        long now = System.currentTimeMillis();
        Iterator<Map.Entry<Path, Long>> it = pending.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<Path, Long> e = it.next();
            if (now >= e.getValue()) {
                it.remove();
                handleFileChange(e.getKey().toFile());
            }
        }
    }

    private void handleFileChange(File changedFile) {
        if (rootFolder == null) return;
        if (!changedFile.exists()) return;
        List<FolderScanner.CompanyFile> companies = scanner.scan(rootFolder);
        Optional<FolderScanner.CompanyFile> match = companies.stream()
            .filter(cf -> cf.excelFile().getAbsolutePath()
                .equals(changedFile.getAbsolutePath()))
            .findFirst();

        if (match.isPresent()) {
            FolderScanner.CompanyFile cf = match.get();
            onEvent.accept(cf.companyName(), "Changement détecté → synchronisation...");
            try {
                syncService.syncCompany(cf);
                onEvent.accept(cf.companyName(), "✓ Synchronisé");
            } catch (Exception ex) {
                onEvent.accept(cf.companyName(), "✗ Erreur : " + ex.getMessage());
            }
        } else {
            onEvent.accept("WATCH", "Fichier modifié (non suivi) : " + changedFile.getName());
        }
    }


    private void pollCompanyFiles() {
        if (rootFolder == null) return;
        try {
            List<FolderScanner.CompanyFile> companies = scanner.scan(rootFolder);
            for (FolderScanner.CompanyFile cf : companies) {
                Path p = cf.excelFile().toPath();
                long current = cf.excelFile().lastModified();
                Long previous = lastModified.get(p);
                if (previous != null && current != previous) {
                    pending.put(p, System.currentTimeMillis() + 2000);
                    onEvent.accept(cf.companyName(), "[POLL] Changement détecté");
                }
                lastModified.put(p, current);
            }
        } catch (Exception ignored) {
            // polling failure is non-fatal
        }
    }

    private void registerAll(Path start, WatchService ws) throws IOException {
        Files.walkFileTree(start, new SimpleFileVisitor<>() {
            @Override
            public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs)
                    throws IOException {
                dir.register(ws, ENTRY_CREATE, ENTRY_MODIFY, ENTRY_DELETE);
                return FileVisitResult.CONTINUE;
            }
        });
    }
}
