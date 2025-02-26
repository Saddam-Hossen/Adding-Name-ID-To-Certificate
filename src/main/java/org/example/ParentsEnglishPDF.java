package org.example;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.example.AddTextsToPdf.deleteAllFilesInFolder;

public class ParentsEnglishPDF {
    public static void main(String[] args) {
        deleteAllFilesInFolder("C:\\reactjs\\ics\\pdf\\singlePdf");
        String excelFilePath = "C:\\reactjs\\ics\\reserve\\Book2.xlsx";

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            int rowCount = sheet.getPhysicalNumberOfRows();
            List<String> names = new ArrayList<>();
            int fileIndex = 1;

            for (int i = 1; i < rowCount; i++) { // Skip the first row (header)
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String name = cell1.toString().trim();

                    if (!name.isEmpty()) {
                        names.add(name);
                    }

                    if (names.size() == 4 || i == rowCount - 1) { // Generate PDF every 4 names or at the end
                        createPDF(names, fileIndex);
                        names.clear();
                        fileIndex++;
                    }
                }
            }

            //
           // generate pdf single
            String inputFolder = "C:\\reactjs\\ics\\pdf\\singlePdf";  // Folder containing PDFs
            String outputPdf = "C:\\reactjs\\ics\\pdf\\Parents_SNVN.pdf";  // Final merged PDF

            mergePdfFiles(inputFolder, outputPdf);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void createPDF(List<String> names, int fileIndex) {
        String inputPdf = "C:\\reactjs\\ics\\reserve\\PAR BOX .pdf";
        String outputPdf = "C:\\reactjs\\ics\\pdf\\singlePdf\\" + fileIndex + ".pdf";

        try {
            PdfReader reader = new PdfReader(inputPdf);
            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(outputPdf));
            document.open();
            PdfContentByte canvas = writer.getDirectContent();
            document.newPage();
            PdfImportedPage page = writer.getImportedPage(reader, 1);
            canvas.addTemplate(page, 0, 0);

            BaseFont calibriBold = BaseFont.createFont("c:/windows/fonts/calibrib.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            String hexColor = "#3456A6";
            // Convert the hex color to a BaseColor object
            int fontSize=8;

            BaseColor color = hexToBaseColor(hexColor);


            int startX = (int) ((document.right() + document.left()) / 2);
            int startY = 500; // Adjusted to fit four names on the page
            int spacing = 50;

            for (int i = 0; i < names.size(); i++) {
                fontSize=12;
                String text = names.get(i).trim();
                int spaceCount = text.length() - text.replace(" ", "").length();
                int totalLength = text.length() + spaceCount;
                System.out.println(text +"  "+totalLength);
                if(totalLength>19 && totalLength<28){
                    fontSize=9;

                }
                else if(totalLength>19 && totalLength<32)
                {
                    fontSize=8;
                }
                else if(totalLength>=32){
                    fontSize=7;
                }
                Font font = new Font(calibriBold, fontSize, Font.BOLD, BaseColor.BLACK);//12
           if(i==0)
           {
               ColumnText.showTextAligned(canvas, Element.ALIGN_CENTER, new Phrase(names.get(i), font), 165, 630 , 0);

           }
           else if(i==1){
             //  ColumnText.showTextAligned(canvas, Element.ALIGN_CENTER, new Phrase(names.get(i), font), startX, startY - (i * spacing), 0);
               ColumnText.showTextAligned(canvas, Element.ALIGN_CENTER, new Phrase(names.get(i), font), 435, 626 , 0);

           }

           else if(i==2){
              // ColumnText.showTextAligned(canvas, Element.ALIGN_CENTER, new Phrase(names.get(i), font), startX, startY - (i * spacing), 0);
               ColumnText.showTextAligned(canvas, Element.ALIGN_CENTER, new Phrase(names.get(i), font), 165, 235 , 0);

           }
           else{
              // ColumnText.showTextAligned(canvas, Element.ALIGN_CENTER, new Phrase(names.get(i), font), startX, startY - (i * spacing), 0);
               ColumnText.showTextAligned(canvas, Element.ALIGN_CENTER, new Phrase(names.get(i), font), 415, 238 , 0);

           }
            }

            document.close();
            writer.close();
            reader.close();

            System.out.println("Generated: " + outputPdf);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void mergePdfFiles(String folderPath, String outputFilePath) {
        File folder = new File(folderPath);
        File[] files = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".pdf"));

        if (files == null || files.length == 0) {
            System.out.println("No PDF files found in the folder.");
            return;
        }

        Arrays.sort(files); // Sort files alphabetically

        try {
            Document document = new Document();
            PdfCopy copy = new PdfCopy(document, new FileOutputStream(outputFilePath));
            document.open();

            for (File file : files) {
                System.out.println("Merging: " + file.getName());
                PdfReader reader = new PdfReader(file.getAbsolutePath());
                int pages = reader.getNumberOfPages();

                for (int i = 1; i <= pages; i++) {
                    copy.addPage(copy.getImportedPage(reader, i));
                }
                reader.close();
            }

            document.close();
            System.out.println("Merged PDF created: " + outputFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    // Helper method to convert a hex color string to a BaseColor object
    public static BaseColor hexToBaseColor(String hexColor) {
        // Remove the '#' if present
        if (hexColor.startsWith("#")) {
            hexColor = hexColor.substring(1);
        }

        // Parse the hex color into RGB values
        int r = Integer.parseInt(hexColor.substring(0, 2), 16);
        int g = Integer.parseInt(hexColor.substring(2, 4), 16);
        int b = Integer.parseInt(hexColor.substring(4, 6), 16);

        return new BaseColor(r, g, b);
    }
}
