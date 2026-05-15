package com.creances;

import java.io.File;
import java.net.URI;
import java.net.URL;

public class TestFixtures {
    public static File get(String name) {
        URL url = TestFixtures.class.getClassLoader().getResource("fixtures/" + name);
        if (url == null) throw new IllegalStateException("Fixture not found: " + name);
        try {
            return new File(url.toURI());
        } catch (Exception e) {
            return new File(url.getFile());
        }
    }

    public static boolean exists(String name) {
        URL url = TestFixtures.class.getClassLoader().getResource("fixtures/" + name);
        return url != null;
    }
}
