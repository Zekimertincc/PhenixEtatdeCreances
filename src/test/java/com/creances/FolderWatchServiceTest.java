package com.creances;

import com.zeki.merger.service.EtatCreancesSyncService;
import com.zeki.merger.service.FolderScanner;
import com.zeki.merger.service.FolderWatchService;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;

import static org.assertj.core.api.Assertions.*;

class FolderWatchServiceTest {

    @Test
    void startsAndStopsCleanly(@TempDir Path tempDir) throws Exception {
        EtatCreancesSyncService sync = new EtatCreancesSyncService(null);
        FolderWatchService svc = new FolderWatchService(sync, (c, m) -> {});

        svc.start(tempDir.toFile());
        assertThat(svc.isRunning()).isTrue();

        svc.stop();
        Thread.sleep(300);
        assertThat(svc.isRunning()).isFalse();
    }

    @Test
    void detectsXlsxChange(@TempDir Path tempDir) throws Exception {
        Path companyDir  = tempDir.resolve("TEST COMPANY");
        Path creancesDir = companyDir.resolve("Etat des créances");
        Files.createDirectories(creancesDir);

        CountDownLatch latch  = new CountDownLatch(1);
        List<String>   events = new ArrayList<>();

        EtatCreancesSyncService sync = new EtatCreancesSyncService(null) {
            @Override
            public void syncCompany(FolderScanner.CompanyFile cf) {
                events.add(cf.companyName());
                latch.countDown();
            }
        };

        FolderWatchService svc = new FolderWatchService(sync, (c, m) -> {});
        svc.start(tempDir.toFile());
        Thread.sleep(500);

        Files.writeString(creancesDir.resolve("etat_test.xlsx"), "dummy");

        boolean triggered = latch.await(6, TimeUnit.SECONDS);
        svc.stop();

        assertThat(triggered).as("WatchService doit détecter le changement").isTrue();
        assertThat(events).anyMatch(e -> e.equalsIgnoreCase("TEST COMPANY"));
    }

    @Test
    void ignoresTempFiles(@TempDir Path tempDir) throws Exception {
        Path companyDir  = tempDir.resolve("TEMP TEST");
        Path creancesDir = companyDir.resolve("Etat des créances");
        Files.createDirectories(creancesDir);

        List<String> synced = new ArrayList<>();
        EtatCreancesSyncService sync = new EtatCreancesSyncService(null) {
            @Override
            public void syncCompany(FolderScanner.CompanyFile cf) {
                synced.add(cf.companyName());
            }
        };

        FolderWatchService svc = new FolderWatchService(sync, (c, m) -> {});
        svc.start(tempDir.toFile());
        Thread.sleep(500);

        Files.writeString(creancesDir.resolve("~$etat_test.xlsx"), "dummy");
        Thread.sleep(3500);

        svc.stop();
        assertThat(synced).isEmpty();
    }
}
