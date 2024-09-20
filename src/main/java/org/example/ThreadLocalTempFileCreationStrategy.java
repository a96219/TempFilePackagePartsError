package org.example;

import org.apache.commons.io.FileUtils;
import org.apache.poi.util.TempFileCreationStrategy;

import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.UUID;

public class ThreadLocalTempFileCreationStrategy implements TempFileCreationStrategy {

    private final static ThreadLocal<File> dir = new ThreadLocal<>();

    private final URI uri;

    public ThreadLocalTempFileCreationStrategy(URI uri) {
        this.uri = uri;
    }

    @Override
    public File createTempFile(String prefix, String suffix) throws IOException {
        return Files.createTempFile(getDir().toPath(), prefix, suffix).toFile();
    }

    @Override
    public File createTempDirectory(String prefix) throws IOException {
        return Files.createTempDirectory(getDir().toPath(), prefix).toFile();
    }

    private File getDir() throws IOException {
        var file = dir.get();
        if (file == null) {
            file = Files.createDirectory(Paths.get(uri.resolve(UUID.randomUUID().toString()))).toFile();
            dir.set(file);
        }

        return file;
    }

    public static void clean() throws IOException {
        var file = dir.get();
        if (file != null) {
            dir.remove();
            FileUtils.deleteDirectory(file);
        }
    }
}
