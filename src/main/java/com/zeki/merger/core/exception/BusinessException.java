package com.zeki.merger.core.exception;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

public class BusinessException extends RuntimeException {

    private final ErrorCode errorCode;
    private final Map<String, Object> context;

    public BusinessException(ErrorCode errorCode, String message) {
        super(message);
        this.errorCode = errorCode;
        this.context = new HashMap<>();
    }

    public BusinessException(ErrorCode errorCode, String message, Throwable cause) {
        super(message, cause);
        this.errorCode = errorCode;
        this.context = new HashMap<>();
    }

    public BusinessException(ErrorCode errorCode, String message, Map<String, Object> context) {
        super(message);
        this.errorCode = errorCode;
        this.context = new HashMap<>(context);
    }

    public ErrorCode getErrorCode() { return errorCode; }

    public Map<String, Object> getContext() {
        return Collections.unmodifiableMap(context);
    }

    public String getDetailedMessage() {
        StringBuilder sb = new StringBuilder()
            .append("[").append(errorCode.getCode()).append("] ")
            .append(getMessage());
        if (!context.isEmpty()) {
            sb.append(" | Context: ").append(context);
        }
        return sb.toString();
    }
}
