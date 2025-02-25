package org.example;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Set;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class AddTextsToPdf {
    static int count=1;
    public static void main(String[] args) {
        String folderPath = "G:/pdf/sodosso"; // Replace with your folder path
        String xlsxFile = "G:\\pdf\\সদস্য_তালিকা-2024.xlsx";
        int k = 0;

        try (FileInputStream fis = new FileInputStream(new File(xlsxFile));
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Get the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Skip the header row and iterate over rows
            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue; // Skip the header row
                }

                // Retrieve the necessary cells
                Cell idCell = row.getCell(0); // Adjust cell indexes based on your data
                Cell nameCell = row.getCell(1);
                Cell additionalInfoCell = row.getCell(2);

                if (idCell != null && nameCell != null && additionalInfoCell != null && k < 50) {
                    k++;

                    // Retrieve values as strings
                    String id = idCell.toString();
                    String name = nameCell.toString();
                    String additionalInfo = additionalInfoCell.toString();

                    // Paths to images and the output file
                    String existingImagePath = "G:\\pdf/banglafont/শাখার সদস্য-04.jpg";
                    String outputImagePath = "G:\\pdf/sodossoFinal/" + name + "_" + id + ".jpg";
                    String outputImg = "G:\\pdf/sodosso/" + id + ".png";
                    String customFontPath = "G:\\pdf/banglafont/Bornomala-Bold.ttf";

                    // Create and overlay images
                    String inputImagePath = createImageFromBengaliText(id, customFontPath, outputImg);
                    String resultPath = overlayImage(existingImagePath, inputImagePath, countWords(id), 1125, outputImagePath);

                    inputImagePath = createImageFromBengaliText(additionalInfo, customFontPath, outputImg);
                    resultPath = overlayImage(resultPath, inputImagePath, countWords(additionalInfo), 1220, outputImagePath);

                    inputImagePath = createImageFromBengaliText(name, customFontPath, outputImg);
                    resultPath = overlayImage(resultPath, inputImagePath, countWords(name), 1320, outputImagePath);

                    // Output path verification
                    System.out.println("Overlay complete. New image saved at: " + resultPath);
                }
            }
        } catch (IOException | FontFormatException e) {
            e.printStackTrace();
        }

        // Combine all images into a PDF
         generateFinalPdf();
        generateBranchPdf();

        // deleteAllFilesInFolder("G:\\pdf/sodossoFinal");
    }
public static void generateBranchPdf(){
    File[] files = getPdfFilesPdf("G:\\pdf/sodossoFinalM");

    if (files == null || files.length == 0) {
        System.out.println("No image files found in the directory2.");
        return;
    }
    Set<String> splitPathSet = new HashSet<>();

    for (File file : files) {
        String absolutePath = file.getAbsolutePath();

        // Split the path by "_"
        String[] parts = absolutePath.split("_");

        // Add the 0 index part (if exists) to the Set
        if (parts.length > 0) {
            splitPathSet.add(parts[0]);
        }
    }

    splitPathSet.forEach(e->{
        ArrayList<String> filenames=new ArrayList();

        for (File file : files) {
            String absolutePath = file.getAbsolutePath();

            // Split the path by "_"
            String[] parts = absolutePath.split("_");

            // Add the 0 index part (if exists) to the Set
            if (parts.length > 0) {
                splitPathSet.add(parts[0]);
                if(e.equals(parts[0])){
                   filenames.add(absolutePath);
                }
            }
        }
        String[] data=new String[filenames.size()];
        int l=0;
        for (String item : filenames) {
            data[l]=item;
            l++;
            System.out.println(item);
        }
        System.out.println(e);
        try {
            String outputPdf =e + ".pdf";
            String regex = "([^\\/\\\\]+)\\.[^\\.]+$"; // Matches everything before the last dot

            Pattern pattern = Pattern.compile(regex);
            Matcher matcher = pattern.matcher(outputPdf);

            if (matcher.find()) {
                String fileNameWithExtension = matcher.group(1) + matcher.group(0).substring(matcher.group(1).length());
                System.out.println("File Name with Extension: " + fileNameWithExtension);
                combineFullPdfs(data, "G:\\pdf/sodossoFinalBranch/"+fileNameWithExtension);
            } else {
                System.out.println("File name extraction failed.");
            }

        } catch (DocumentException ex) {
            throw new RuntimeException(ex);
        } catch (IOException ex) {
            throw new RuntimeException(ex);
        }

    });


}
    public static int countWords(String input) {
        if (input == null) {
            return 0; // Return 0 for null strings
        }

        return 400; // Use the length() method to count characters
    }
    public static void generateFinalPdf() {
        // Fetch all image files from the given folder
        File[] files = getPdfFiles("G:\\pdf/sodossoFinal");

        if (files == null || files.length == 0) {
            System.out.println("No image files found in the directory.");
            return;
        }
        // Store split file path parts in a Set
        Set<String> splitPathSet = new HashSet<>();

        for (File file : files) {
            String absolutePath = file.getAbsolutePath();

            // Split the path by "_"
            String[] parts = absolutePath.split("_");

            // Add the 0 index part (if exists) to the Set
            if (parts.length > 0) {
                splitPathSet.add(parts[0]);
            }
        }
        String outputFolderPath = "G:\\pdf/sodossoFinalM/";
        AtomicInteger fileCounter = new AtomicInteger(); // Counter to manage output file names
        // Example: Using splitPathSet for further file filtering/processing
        splitPathSet.forEach(prefix -> {

            // Filter files matching the current prefix
            File[] filteredFiles = filterFilesByPrefix(files, prefix);

            // Process filtered files as needed
            System.out.println("Files with prefix: " + prefix);
            for (File filteredFile : filteredFiles) {
                System.out.println("  - " + filteredFile.getAbsolutePath());
            }

            count=1;

            // Process the files in groups of four
            for (int i = 0; i < filteredFiles.length; i += 4) {


                String[] imageData = new String[4];

                // Add images to the group
                for (int j = 0; j < 4; j++) {
                    if (i + j < filteredFiles.length) {
                        imageData[j] = filteredFiles[i + j].getAbsolutePath(); // Use available images
                    } else {
                        // Repeat images if fewer than 4 are left
                        imageData[j] = filteredFiles[filteredFiles.length - 1].getAbsolutePath();
                    }
                }

                // Call createPDFWithCombinedImages for the current group of images
                try {
                    String[] pathParts = filteredFiles[0].getAbsolutePath().split("\\\\");
                    String filename = pathParts[pathParts.length - 1]; // Get the name part only

                    String outputPdfPath = outputFolderPath + filename.split("_")[0]+"_" + (count) + ".pdf";  // Use prefix only as filename
                    createPdfFromImages(imageData, outputPdfPath);
                    System.out.println("Generated: " + outputPdfPath);
                    count++;
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });



    }


    // Method to filter files by prefix
    public static File[] filterFilesByPrefix(File[] files, String prefix) {
        return java.util.Arrays.stream(files)
                .filter(file -> file.getAbsolutePath().startsWith(prefix))
                .toArray(File[]::new);
    }
    public static String overlayImage(String existingImagePath, String inputImagePath, int xPosition, int yPosition, String outputImagePath) throws IOException {
        // Read the existing image and input image
        BufferedImage existingImage = ImageIO.read(new File(existingImagePath));
        BufferedImage inputImage = ImageIO.read(new File(inputImagePath));

        // Check if the images are read properly
        if (existingImage == null) {
            System.out.println("Failed to read the existing image from path: " + existingImagePath);
            return null;
        }

        if (inputImage == null) {
            System.out.println("Failed to read the input image from path: " + inputImagePath);
            return null;
        }

        // Debugging: Image dimensions
        System.out.println("Existing image size: " + existingImage.getWidth() + "x" + existingImage.getHeight());
        System.out.println("Input image size: " + inputImage.getWidth() + "x" + inputImage.getHeight());

        // Create a new buffered image with the same dimensions as the existing image (with transparency support)
        BufferedImage combinedImage = new BufferedImage(
                existingImage.getWidth(),
                existingImage.getHeight(),
                BufferedImage.TYPE_INT_ARGB  // Changed to keep transparency
        );

        // Create a Graphics2D object
        Graphics2D g2d = combinedImage.createGraphics();

        // Draw the existing image and the input image onto the new image
        g2d.drawImage(existingImage, 0, 0, null);

        // Ensure input image is within bounds
        if (xPosition >= 0 && yPosition >= 0 &&
                xPosition + inputImage.getWidth() <= existingImage.getWidth() &&
                yPosition + inputImage.getHeight() <= existingImage.getHeight()) {

            g2d.drawImage(inputImage, xPosition, yPosition, null);
        } else {
            System.out.println("Input image out of bounds, skipping overlay.");
        }

        g2d.dispose();

        // Check if the output directory exists
        File outputFile = new File(outputImagePath);
        if (outputFile.getParentFile() != null && !outputFile.getParentFile().exists()) {
            System.out.println("Directory does not exist: " + outputFile.getParent());
            return null;
        } else {
            System.out.println("Output image path: " + outputImagePath);
        }

        // Write the resulting image to the output file
        try {
            ImageIO.write(combinedImage, "png", new File(outputImagePath)); // Save as PNG to keep transparency
            System.out.println("Image written successfully at: " + outputImagePath);
        } catch (IOException e) {
            System.out.println("Error writing the image: " + e.getMessage());
            return null;
        }

        // Return the output image path
        return outputImagePath;
    }



    // Method to create an image from Bengali text and save it as a PNG
    public static String createImageFromBengaliText(String bengaliText, String fontPath, String outputImagePath) throws IOException, FontFormatException {
        // Load the custom font
        Font font = Font.createFont(Font.TRUETYPE_FONT, new FileInputStream(fontPath));
        font = font.deriveFont(50f); // Adjust the font size as needed

        // Create a BufferedImage for drawing
        BufferedImage bufferedImage = new BufferedImage(600, 200, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics2D = bufferedImage.createGraphics();

        // Set rendering quality and background color
        graphics2D.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        graphics2D.setColor(Color.WHITE); // Background color

        // Set the text color and font
        graphics2D.setColor(Color.BLACK);
        graphics2D.setFont(font);

        // Calculate the width and height of the text to center it on the image
        int textWidth = graphics2D.getFontMetrics().stringWidth(bengaliText);
        int textHeight = graphics2D.getFontMetrics().getHeight();

        // Center the text
        int x = (bufferedImage.getWidth() - textWidth) / 2;
        int y = (bufferedImage.getHeight() + textHeight) / 2;

        // Draw the Bengali text onto the image
        graphics2D.drawString(bengaliText, x, y);

        // Dispose of the graphics context
        graphics2D.dispose();

        // Save the image to the specified file path
        File outputFile = new File(outputImagePath);
        ImageIO.write(bufferedImage, "PNG", outputFile);

        // Return the image path
        return outputFile.getAbsolutePath();
    }

    public static void deleteAllFilesInFolder(String folderPath) {
        File folder = new File(folderPath);

        if (!folder.exists() || !folder.isDirectory()) {
            System.out.println("Invalid folder path: " + folderPath);
            return;
        }

        File[] files = folder.listFiles();

        if (files == null || files.length == 0) {
            System.out.println("The folder is already empty.");
            return;
        }

        for (File file : files) {
            if (file.isFile()) {
                if (file.delete()) {
                    System.out.println("Deleted file: " + file.getName());
                } else {
                    System.err.println("Failed to delete file: " + file.getName());
                }
            }
        }

        System.out.println("All files in the folder have been removed.");
    }

    public static File[] getPdfFiles(String folderPath) {
        File folder = new File(folderPath);
        return folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".jpg"));
    }
    public static File[] getPdfFilesPdf(String folderPath) {
        File folder = new File(folderPath);
        return folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".pdf"));
    }

    public static void combineFullPdfs(String[] inputPdfFiles, String outputPdfFile) throws DocumentException, IOException {
        if (inputPdfFiles == null || inputPdfFiles.length == 0) {
            throw new IllegalArgumentException("You must provide at least one PDF file.");
        }

        // Create a new document with A4 page size
        Document document = new Document(PageSize.A4);
        PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(outputPdfFile));

        document.open();

        // Access the PdfContentByte layer to draw each imported page
        PdfContentByte canvas = writer.getDirectContent();

        // Process each input PDF
        for (String inputPdfFile : inputPdfFiles) {
            PdfReader reader = new PdfReader(inputPdfFile);
            PdfImportedPage importedPage = writer.getImportedPage(reader, 1);

            // Get the dimensions of the imported page
            float pageWidth = importedPage.getWidth();
            float pageHeight = importedPage.getHeight();

            // Scale the imported page to fit within the output page
            float scaleX = document.getPageSize().getWidth() / pageWidth;
            float scaleY = document.getPageSize().getHeight() / pageHeight;
            float scale = Math.min(scaleX, scaleY);

            // Calculate x and y offsets to center the imported page
            float xOffset = (document.getPageSize().getWidth() - pageWidth * scale) / 2;
            float yOffset = (document.getPageSize().getHeight() - pageHeight * scale) / 2;

            // Add the imported page to the canvas, scaling and positioning it
            canvas.addTemplate(importedPage, scale, 0, 0, scale, xOffset, yOffset);

            // Start a new page for the next PDF
            document.newPage();
        }

        document.close();
        System.out.println("Combined PDF created successfully: " + outputPdfFile);
    }
    // Method that creates a PDF from four given images arranged in a 2x2 grid with a 20px outer margin
    public static void createPdfFromImages(String[] imagePaths, String outputPdfPath) throws Exception {
        // A4 page dimensions in points (1 inch = 72 points, A4 = 8.27in x 11.69in)
        com.itextpdf.text.Rectangle pageSize = new com.itextpdf.text.Rectangle(595, 842);  // A4 size in points
        float outerMargin = 20f;  // Outer margin around the grid in pixels (20px)

        // Create a new PDF document with A4 dimensions and an outer margin
        Document document = new Document(pageSize);

        // Create PdfWriter instance to write the document into a file
        PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(outputPdfPath));
        document.open();

        // Define image size for each image (you can adjust the size here)
        int imageWidth = 255;   // Width of each image in points
        int imageHeight = 380;  // Height of each image in points

        // Define custom gaps for each image in the grid
        int gap1 = 20; // Gap between image 1 and image 2 (horizontal)
        int gap2 = 20; // Gap between image 3 and image 4 (horizontal)
        int gap3 = 20; // Gap between image 1 and image 3 (vertical)
        int gap4 = 20; // Gap between image 2 and image 4 (vertical)

        // Load the images from the provided file paths
        com.itextpdf.text.Image[] pdfImages = new com.itextpdf.text.Image[imagePaths.length];
        for (int i = 0; i < imagePaths.length; i++) {
            pdfImages[i] = com.itextpdf.text.Image.getInstance(imagePaths[i]);
            pdfImages[i].scaleToFit(imageWidth, imageHeight);  // Scale the image to the desired size
        }

        // Calculate the total width and height considering the images, gaps, and outer margin
        float totalWidth = imageWidth * 2 + gap1 + gap2 + 2 * outerMargin;
        float totalHeight = imageHeight * 2 + gap3 + gap4 + 2 * outerMargin;

        // Positioning images inside the outer margin, center them on the page
        float xOffset = (pageSize.getWidth() - totalWidth) / 2;
        float yOffset = (pageSize.getHeight() - totalHeight) / 2;

        // Position of each image with the outer margin and custom gaps
        pdfImages[0].setAbsolutePosition(xOffset + outerMargin, yOffset + outerMargin + imageHeight + gap3 + 25);  // Image 1
        pdfImages[1].setAbsolutePosition(xOffset + outerMargin + imageWidth + gap1+25, yOffset + outerMargin + imageHeight + gap3 + 25);  // Image 2
        pdfImages[2].setAbsolutePosition(xOffset + outerMargin, yOffset + outerMargin);  // Image 3
        pdfImages[3].setAbsolutePosition(xOffset + outerMargin + imageWidth + gap2+25, yOffset + outerMargin);  // Image 4

        // Add the images to the PDF document
        for (Image img : pdfImages) {
            document.add(img);
        }

        // Close the document
        document.close();

        System.out.println("PDF created successfully at: " + outputPdfPath);
    }

}






