package org.example;

import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.opc.ZipPackage;
import org.apache.poi.openxml4j.util.ZipInputStreamZipEntrySource;
import org.apache.poi.util.DefaultTempFileCreationStrategy;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Objects;

public class Main {

    public static void main(String[] args) throws IOException {
        ZipInputStreamZipEntrySource.setThresholdBytesForTempFiles(50000000);
        ZipPackage.setUseTempFilePackageParts(true);

        var path = Objects.requireNonNull(Main.class.getClassLoader().getResource("")).getPath();
        var tempPath = path + "/temp";
        var tempDir = new File(tempPath);
        if (!tempDir.exists()) {
            if (!tempDir.mkdir()) {
                throw new RuntimeException("Could not create directory: " + tempDir);
            }
        } else {
            FileUtils.deleteDirectory(tempDir);
            if (!tempDir.mkdir()) {
                throw new RuntimeException("Could not create directory: " + tempDir);
            }
        }
        TempFile.setTempFileCreationStrategy(new DefaultTempFileCreationStrategy(tempDir));

        try (var wb = new SXSSFWorkbook(1)) {
            var sheet = wb.createSheet("test");
            sheet.setColumnWidth(3, 5000);
            sheet.setDefaultRowHeightInPoints(100F);

            var drawingPatriarch = sheet.createDrawingPatriarch();
            for (var i = 0; i < 200; i++) {
                var row = sheet.createRow(i);

                row.createCell(0).setCellValue(i + "");
                row.createCell(1).setCellValue(i + "");
                row.createCell(2).setCellValue(i + "");

                try (var inputStream = Main.class.getClassLoader().getResourceAsStream("test.jpg")) {
                    byte[] bytes = null;
                    if (inputStream != null) {
                        bytes = inputStream.readAllBytes();
                    }

                    var clientAnchor = new XSSFClientAnchor(0, 0, 0, 0, 3, i, 4, i + 1);
                    drawingPatriarch.createPicture(clientAnchor, wb.addPicture(bytes, SXSSFWorkbook.PICTURE_TYPE_JPEG));
                    System.out.println(i);
                }
            }

            try (var file = new FileOutputStream(path + "/text.xlsx")) {
                wb.write(file);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}