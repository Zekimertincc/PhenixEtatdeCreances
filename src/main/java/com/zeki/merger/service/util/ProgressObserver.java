package com.zeki.merger.service.util;

public interface ProgressObserver {
    void onProgressUpdate(double progress, String message);
    void onCompleted();
    void onFailed(Exception exception);
}
