package org.example;

import java.io.File;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashSet;
import java.util.Set;

public class Main {
    public static void main(String[] args) throws Exception {
        SheetParser parser = new SheetParser();
        Set<String> fileSet = new HashSet<>();

        Path directoryPath = Paths.get(System.getProperty("user.dir")).resolve("documents");

        try (DirectoryStream stream = Files.newDirectoryStream(directoryPath)) {
            int index = 1;
            for (Object path : stream) {
                if (Files.isRegularFile((Path) path) && isExcelFile((Path) path))
                    fileSet.add(directoryPath + File.separator + ((Path) path).getFileName().toString());
            }
            fileSet.forEach(value ->
                    System.out.println(" File: " + value));
        } catch (IOException e) {
            System.err.println("Error reading directory: " + e.getMessage());
        }
        parser.matchingRules();
        fileSet.forEach(parser::parse);
        parser.save();
    }

    private static boolean isExcelFile(Path path) {
        String fileName = path.getFileName().toString().toLowerCase();
        return fileName.endsWith(".xls") || fileName.endsWith(".xlsx");
    }

}
