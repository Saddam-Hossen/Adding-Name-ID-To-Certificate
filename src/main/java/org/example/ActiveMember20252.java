
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Set;

public class ActiveMember20252 {
    static int count=1;
    public static void main(String[] args) throws DocumentException, IOException {
        deleteAllFilesInFolder("D:\\NameToImageForDeligete\\sodossoFinalM");
        deleteAllFilesInFolder("D:\\NameToImageForDeligete\\sodossoFinalBranch");
        deleteAllFilesInFolder("D:\\NameToImageForDeligete\\sodossoFinal");
        deleteAllFilesInFolder("D:\\NameToImageForDeligete\\sodosso");

        String folderPath = "D:\\NameToImageForDeligete\\sodosso"; // Replace with your folder path
        // String xlsxFile = "G:\\pdf\\Delegate Card Update.xlsx";// for two
        String xlsxFile = "D:\\NameToImageForDeligete\\পরিষদ বৈঠকের ডেলিগট_2025.xlsx";// for three
        File excelFile = new File(xlsxFile);

        // Check if the file exists
        if (!excelFile.exists()) {
            System.out.println("File does not exist at: " + xlsxFile);
            return;
        }

        // Get the file size in bytes
        long fileSizeInBytes = excelFile.length();

        // Convert the file size to kilobytes (KB) and megabytes (MB)
        double fileSizeInKB = fileSizeInBytes / 1024.0;
        double fileSizeInMB = fileSizeInKB / 1024.0;

        // Display the file size
        System.out.println("File Size:");
        System.out.println("Bytes: " + fileSizeInBytes);
        System.out.println("Kilobytes: " + String.format("%.2f", fileSizeInKB));
        System.out.println("Megabytes: " + String.format("%.2f", fileSizeInMB));
        int k = 1;

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
                Cell nameCell = row.getCell(2);// age 3
                Cell additionalInfoCell = row.getCell(1);// age 3
                // Cell uniqueIdC= row.getCell(0);

                if (idCell != null && nameCell != null && additionalInfoCell != null && k < 100) {
                    k++;
                    //  String uniqueId=uniqueIdC.toString();
                    // Retrieve values as strings
                    String id = idCell.toString();
                    String name = nameCell.toString();
                    String additionalInfo = additionalInfoCell.toString();

                    // Paths to images and the output file
                    // String existingImagePath = "G:\\pdf/banglafont/Delegate2.jpg";// for two
                    String existingImagePath = "D:\\NameToImageForDeligete\\reserve/কার্যকরী_পরিষদের_২য়_সাধারণ_অধিবেশন_25_ID_card_01 (2).jpg";// for three
                    String outputImagePath = "D:\\NameToImageForDeligete\\sodossoFinal/"+k+"_" + name + ".jpg";
                    String outputImg = "D:\\NameToImageForDeligete\\sodosso/" + k + ".png";
                    String customFontPath = "D:\\NameToImageForDeligete\\reserve/SolaimanLipi_Bold_10-03-12.ttf";

                    // Create and overlay images
                   // String inputImagePath = createImageFromBengaliText((id), customFontPath, outputImg,"1");
                    //String resultPath = overlayImage(existingImagePath, inputImagePath, 390, 1040, outputImagePath);

                    String inputImagePath = createImageFromBengaliText(additionalInfo, customFontPath, outputImg,"1");
                    String resultPath = overlayImage(existingImagePath, inputImagePath, 360, 1135, outputImagePath);

                    inputImagePath = createImageFromBengaliText((name), customFontPath, outputImg,"2");
                    resultPath = overlayImage(resultPath, inputImagePath, 360, 1235, outputImagePath);

                    // Output path verification
                    System.out.println("Overlay Image created : ("+k +") "+ resultPath);
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
    // Method to validate and sanitize a file name
    public static String getValidFileName(String name) {
        if (name == null || name.isEmpty()) {
            throw new IllegalArgumentException("Input name cannot be null or empty.");
        }
        // Replace invalid characters with an underscore
        return name.replaceAll("[\\\\/:*?\"<>|]", "_");
    }

    public static void generateBranchPdf() throws DocumentException, IOException {
        File[] files = getPdfFilesPdf("D:\\NameToImageForDeligete/sodossoFinalM");

        if (files == null || files.length == 0) {
            System.out.println("No image files found in the directory2.");
            return;
        }
        Set<String> splitPathSet = new HashSet<>();
        ArrayList<String> filenames=new ArrayList();

        for (File file : files) {
            String absolutePath = file.getAbsolutePath();
            filenames.add(absolutePath);

        }
        String[] data=new String[filenames.size()];
        int l=0;
        for (String item : filenames) {
            data[l]=item;
            l++;
            System.out.println(item);
        }

        combineFullPdfs(data, "D:\\NameToImageForDeligete/sodossoFinalBranch/"+"Present Departmental Member.SNVN Ltd.pdf");


    }
    public static int countWords(String input) {
        if (input == null) {
            return 0; // Return 0 for null strings
        }

        return 390; // Use the length() method to count characters
    }
    public static void generateFinalPdf() {
        // Fetch all image files from the given folder
        File[] files = getPdfFiles("D:\\NameToImageForDeligete/sodossoFinal");

        if (files == null || files.length == 0) {
            System.out.println("No image files found in the directory.");
            return;
        }

        String outputFolderPath = "D:\\NameToImageForDeligete/sodossoFinalM/";
        count=1;

        // Process the files in groups of four
        for (int i = 0; i < files.length; i += 4) {


            String[] imageData = new String[4];

            // Add images to the group
            for (int j = 0; j < 4; j++) {
                if (i + j < files.length) {
                    imageData[j] = files[i + j].getAbsolutePath(); // Use available images
                } else {
                    // Repeat images if fewer than 4 are left
                    imageData[j] = files[files.length - 1].getAbsolutePath();
                }
            }

            // Call createPDFWithCombinedImages for the current group of images
            try {
                String[] pathParts = files[0].getAbsolutePath().split("\\\\");
                String filename = pathParts[pathParts.length - 1]; // Get the name part only

                String outputPdfPath = outputFolderPath + filename.split("_")[0]+"_" + (count) + ".pdf";  // Use prefix only as filename
                createPdfFromImages(imageData, outputPdfPath);
                //System.out.println("Generated: " + outputPdfPath);
                count++;
            } catch (Exception e) {
                e.printStackTrace();
            }
        }


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
        // System.out.println("Existing image size: " + existingImage.getWidth() + "x" + existingImage.getHeight());
        // System.out.println("Input image size: " + inputImage.getWidth() + "x" + inputImage.getHeight());

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
            // System.out.println("Output image path: " + outputImagePath);
        }

        // Write the resulting image to the output file
        try {
            ImageIO.write(combinedImage, "png", new File(outputImagePath)); // Save as PNG to keep transparency
            // System.out.println("Image written successfully at: " + outputImagePath);
        } catch (IOException e) {
            System.out.println("Error writing the image: " + e.getMessage());
            return null;
        }

        // Return the output image path
        return outputImagePath;
    }



    // Method to create an image from Bengali text and save it as a PNG
    public static String createImageFromBengaliText(String bengaliText, String fontPath, String outputImagePath,String color) throws IOException, FontFormatException {
        // Load the custom font
        Font font = Font.createFont(Font.TRUETYPE_FONT, new FileInputStream(fontPath));


        // Create a BufferedImage for drawing
        BufferedImage bufferedImage = new BufferedImage(600, 200, BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics2D = bufferedImage.createGraphics();

        // Set rendering quality and background color
        graphics2D.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        graphics2D.setColor(Color.WHITE); // Background color

        // Set the text color and font
        if(color.equals("1")){
            System.out.println("Header 1");
            // Set custom color using the hexadecimal value
            Color customGreen = Color.decode("#000080");
            graphics2D.setColor(customGreen);
            font = font.deriveFont(50f); // Adjust the font size as needed
            // Set the font style to bold
            font = font.deriveFont(Font.BOLD); // Make the font bold

        }
        else{
            graphics2D.setColor(Color.BLACK);
            font = font.deriveFont(40f); // Adjust the font size as needed
        }

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
        int imageWidth = 295;   // Width of each image in points
        int imageHeight = 410;  // Height of each image in points

        // Define custom gaps for each image in the grid
        int gap1 = 3; // Gap between image 1 and image 2 (horizontal)
        int gap2 = 3; // Gap between image 3 and image 4 (horizontal)
        int gap3 = 5; // Gap between image 1 and image 3 (vertical)
        int gap4 = 5; // Gap between image 2 and image 4 (vertical)

        // Load the images from the provided file paths
        Image[] pdfImages = new Image[imagePaths.length];
        for (int i = 0; i < imagePaths.length; i++) {
            pdfImages[i] = Image.getInstance(imagePaths[i]);
            pdfImages[i].scaleToFit(imageWidth, imageHeight);  // Scale the image to the desired size
        }

        // Calculate the total width and height considering the images, gaps, and outer margin
        float totalWidth = imageWidth * 2 + gap1 + gap2 + 2 * outerMargin;
        float totalHeight = imageHeight * 2 + gap3 + gap4 + 2 * outerMargin;

        // Positioning images inside the outer margin, center them on the page
        float xOffset = (pageSize.getWidth() - totalWidth) / 2;
        float yOffset = (pageSize.getHeight() - totalHeight) / 2;

        // Position of each image with the outer margin and custom gaps
        pdfImages[0].setAbsolutePosition(xOffset + outerMargin+5, yOffset + outerMargin + imageHeight + gap3+5 );  // Image 1
        pdfImages[1].setAbsolutePosition(xOffset + outerMargin + imageWidth + gap1+7, yOffset + outerMargin + imageHeight + gap3+5 );  // Image 2
        pdfImages[2].setAbsolutePosition(xOffset + outerMargin+5, yOffset + outerMargin);  // Image 3
        pdfImages[3].setAbsolutePosition(xOffset + outerMargin + imageWidth + gap2+7, yOffset + outerMargin);  // Image 4

        // Add the images to the PDF document
        for (Image img : pdfImages) {
            document.add(img);
        }

        // Close the document
        document.close();

        // System.out.println("PDF created successfully at: " + outputPdfPath);
    }
    // Method that creates a PDF from four given images arranged in a 2x2 grid with a 20px outer margin
    public static void createPdfFromImages1(String[] imagePaths, String outputPdfPath) throws Exception {
        // A4 page dimensions in points (1 inch = 72 points, A4 = 8.27in x 11.69in)
        com.itextpdf.text.Rectangle pageSize = new com.itextpdf.text.Rectangle(595, 842);  // A4 size in points
        float outerMargin = 20f;  // Outer margin around the grid in pixels (20px)

        // Create a new PDF document with A4 dimensions and an outer margin
        Document document = new Document(pageSize);

        // Create PdfWriter instance to write the document into a file
        PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(outputPdfPath));
        document.open();

        // Define image size for each image (you can adjust the size here)
        int imageWidth = 290;   // Width of each image in points
        int imageHeight = 420;  // Height of each image in points

        // Define custom gaps for each image in the grid
        int gap1 = 6; // Gap between image 1 and image 2 (horizontal)
        int gap2 = 6; // Gap between image 3 and image 4 (horizontal)
        int gap3 = 5; // Gap between image 1 and image 3 (vertical)
        int gap4 = 5; // Gap between image 2 and image 4 (vertical)

        // Load the images from the provided file paths
        Image[] pdfImages = new Image[imagePaths.length];
        for (int i = 0; i < imagePaths.length; i++) {
            pdfImages[i] = Image.getInstance(imagePaths[i]);
            pdfImages[i].scaleToFit(imageWidth, imageHeight);  // Scale the image to the desired size
        }

        // Calculate the total width and height considering the images, gaps, and outer margin
        float totalWidth = imageWidth * 2 + gap1 + gap2 + 2 * outerMargin;
        float totalHeight = imageHeight * 2 + gap3 + gap4 + 2 * outerMargin;

        // Positioning images inside the outer margin, center them on the page
        float xOffset = (pageSize.getWidth() - totalWidth) / 2;
        float yOffset = (pageSize.getHeight() - totalHeight) / 2;

        // Position of each image with the outer margin and custom gaps
        pdfImages[0].setAbsolutePosition(xOffset + outerMargin, yOffset + outerMargin + imageHeight + gap3 );  // Image 1
        pdfImages[1].setAbsolutePosition(xOffset + outerMargin + imageWidth + gap1, yOffset + outerMargin + imageHeight + gap3 );  // Image 2
        pdfImages[2].setAbsolutePosition(xOffset + outerMargin, yOffset + outerMargin);  // Image 3
        pdfImages[3].setAbsolutePosition(xOffset + outerMargin + imageWidth + gap2, yOffset + outerMargin);  // Image 4

        // Add the images to the PDF document
        for (Image img : pdfImages) {
            document.add(img);
        }

        // Close the document
        document.close();

        // System.out.println("PDF created successfully at: " + outputPdfPath);
    }
}