package com.zeki.merger.service.util;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.BiConsumer;

/**
 * Observer-pattern notifier. Bridges the existing BiConsumer<Double,String> callback style
 * used throughout the codebase with the new ProgressObserver interface.
 *
 * Usage with legacy BiConsumer:
 *   notifier.asBiConsumer()  →  (prog, msg) -> ...
 */
public class ProgressNotifier {

    private final List<ProgressObserver> observers = new ArrayList<>();

    public void subscribe(ProgressObserver observer) {
        Objects.requireNonNull(observer, "Observer cannot be null");
        observers.add(observer);
    }

    public void unsubscribe(ProgressObserver observer) {
        observers.remove(observer);
    }

    public void notifyProgress(double progress, String message) {
        for (ProgressObserver observer : new ArrayList<>(observers)) {
            try {
                observer.onProgressUpdate(progress, message);
            } catch (Exception e) {
                System.err.println("Observer notification failed: " + e.getMessage());
            }
        }
    }

    public void notifyCompleted() {
        for (ProgressObserver observer : new ArrayList<>(observers)) {
            try {
                observer.onCompleted();
            } catch (Exception e) {
                System.err.println("Observer notification failed: " + e.getMessage());
            }
        }
    }

    public void notifyFailed(Exception exception) {
        for (ProgressObserver observer : new ArrayList<>(observers)) {
            try {
                observer.onFailed(exception);
            } catch (Exception e) {
                System.err.println("Observer notification failed: " + e.getMessage());
            }
        }
    }

    /** Returns a BiConsumer adapter compatible with the legacy service API. */
    public BiConsumer<Double, String> asBiConsumer() {
        return this::notifyProgress;
    }
}
