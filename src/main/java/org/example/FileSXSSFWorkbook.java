package org.example;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;
import java.nio.file.Files;
import java.nio.file.Path;

public class FileSXSSFWorkbook extends SXSSFWorkbook {

    private final OPCPackage opcPackage;

    private final Path path;

    private FileSXSSFWorkbook(OPCPackage opcPackage, URI uri) throws IOException {
        super(new XSSFWorkbook(opcPackage));
        this.opcPackage = opcPackage;
        this.path = Files.createFile(Path.of(uri));
    }

    public FileSXSSFWorkbook(URI uri) throws InvalidFormatException, IOException {
        this(OPCPackage.open(createTempFile()), uri);
    }

    public void write() throws IOException {
        try (var out = new FileOutputStream(path.toFile())) {
            write(out);
        }
    }

    @Override
    public void close() throws IOException {
        super.close();
        opcPackage.close();
        ThreadLocalTempFileCreationStrategy.clean();
    }

    public static File createTempFile() throws IOException {
        var tempFile = TempFile.createTempFile("file-sxssf-template", ".xlsx");
        try (var outputStream = Files.newOutputStream(tempFile.toPath());
             var wb = new XSSFWorkbook()) {
            wb.write(outputStream);
        }

        return tempFile;
    }
}
