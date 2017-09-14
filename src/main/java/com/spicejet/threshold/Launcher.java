package com.spicejet.threshold;

import java.io.File;
import java.io.IOException;

public class Launcher {

    public static void main(String[] args) {
        ExcelReaderUtil reader = new ExcelReaderUtil();
        try {
            int columnNumBase;
            if (args.length > 0) {
                columnNumBase = Integer.parseInt(args[0]);
            } else {
                columnNumBase = 5;
            }
            File directory = new File(Constants.INPUT_DIR);
            File[] fList = directory.listFiles();
            if ((fList != null) && (fList.length > 0)) {
                for (File file : fList) {
                    if (file.getName().endsWith(".xlsx")) {
                        String inputExcelFilePath = Constants.INPUT_DIR + Constants.SEPERATOR + file.getName();
                        String outputExcelFilePath = Constants.OUTPUT_DIR + Constants.SEPERATOR + file.getName();
                        reader.writeBooksFromExcelFile(inputExcelFilePath, outputExcelFilePath,
                                reader.readBooksFromExcelFile(inputExcelFilePath, columnNumBase));
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
