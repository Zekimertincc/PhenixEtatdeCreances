package com.zeki.merger.core.exception;

public enum ErrorCode {
    FILE_NOT_FOUND(1001, "Dosya bulunamadı"),
    INVALID_FORMAT(1002, "Geçersiz dosya formatı"),
    DATABASE_ERROR(1003, "Veritabanı hatası"),
    GENERATION_FAILED(1004, "Rapor oluşturma başarısız"),
    INVALID_DATA(1005, "Geçersiz veri"),
    IO_ERROR(1006, "Giriş/çıkış hatası"),
    NORMALIZATION_ERROR(1007, "Veri normalleştirme hatası"),
    UNKNOWN_COMMAND(1008, "Bilinmeyen komut");

    private final int code;
    private final String defaultMessage;

    ErrorCode(int code, String defaultMessage) {
        this.code = code;
        this.defaultMessage = defaultMessage;
    }

    public int getCode() { return code; }
    public String getDefaultMessage() { return defaultMessage; }
}
