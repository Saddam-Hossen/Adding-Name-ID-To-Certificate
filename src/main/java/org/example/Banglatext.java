package org.example;

import com.itextpdf.text.*;
import com.itextpdf.text.Image;
import com.itextpdf.text.pdf.*;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.Font;
import java.awt.image.BufferedImage;
import java.io.*;

public class Banglatext {

    public static void main(String[] args) {
        String csvFile = "G:\\pdf/sodosso2.csv";
        String line;
        String csvSplitBy = ",";
        int k = 0;

        try (BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(csvFile), "UTF-8"))) {
            br.readLine(); // Skip the header row

            while ((line = br.readLine()) != null) {
                String[] data = line.split(csvSplitBy);

                if (data.length >= 3 ) {
                    k++;

                    String inputPdf = "G:\\pdf/banglafont/ss.pdf";
                    String outputPdf = "G:\\pdf/sodosso/" + data[0] + ".pdf";
                    String outputPdfLast = "G:\\pdf/sodossoFinal/" + data[0] + ".pdf";
                    String outputImg = "G:\\pdf/sodosso/" + data[0] + ".png";
                    String customFontPath = "G:\\pdf/banglafont/Bornomala-Bold.ttf"; // Path to the Bengali font

                    // Process the first column
                    String imagePath = createImageFromBengaliText(data[0], customFontPath, outputImg);
                    String firstPdf = addImageToExistingPdf(inputPdf, outputPdf, imagePath, 105, 150);
                    System.out.println(firstPdf);
                    // Process the second column (use the first PDF as input)
                    imagePath = createImageFromBengaliText(data[2], customFontPath, outputImg);
                    String secondPdf = addImageToExistingPdf1(firstPdf, outputPdf, imagePath, 105, 100);
                    System.out.println(secondPdf);
                    imagePath = createImageFromBengaliText(data[1], customFontPath, outputImg);
                    String finalPdf = addImageToExistingPdf2(secondPdf, outputPdfLast, imagePath, 105, 50);
                    // remove all file in folder

                }
            }

        } catch (IOException | FontFormatException e) {
            e.printStackTrace();
        }
        String folderPath = "G:/pdf/sodosso"; // Replace with your folder path
        deleteAllFilesInFolder(folderPath);
    }

    // Method to create an image from Bengali text and save it as a PNG
    public static String createImageFromBengaliText(String bengaliText, String fontPath, String outputImagePath) throws IOException, FontFormatException {
        // Load the custom font
        Font font = Font.createFont(Font.TRUETYPE_FONT, new FileInputStream(fontPath));
        font = font.deriveFont(20f); // Adjust the font size as needed

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

    public static String addImageToExistingPdf(String existingPdfPath, String outputPdfPath, String generatedImagePath, float xPosition, float yPosition) {
        try {
            // Read the existing PDF
            PdfReader reader = new PdfReader(existingPdfPath);

            // Create a FileOutputStream for the final output file
            try (FileOutputStream outputStream = new FileOutputStream(outputPdfPath)) {
                PdfStamper stamper = new PdfStamper(reader, outputStream);

                // Get the canvas to modify the first page
                PdfContentByte canvas = stamper.getOverContent(1); // Modify page 1
                Image image = Image.getInstance(generatedImagePath);
                image.scaleToFit(500, 300); // Scale image
                image.setAbsolutePosition(xPosition, yPosition); // Set image position
                canvas.addImage(image);

                stamper.close(); // Save changes directly to the output file
            }
            reader.close();

            System.out.println("Image added successfully to the PDF at: " + outputPdfPath);
            return outputPdfPath;

        } catch (Exception e) {
            e.printStackTrace();
            return null; // In case of error
        }
    }

    public static String addImageToExistingPdf1(String existingPdfPath, String outputPdfPath, String generatedImagePath, float xPosition, float yPosition) {
        try {
            // Use PdfReader and PdfStamper for modification
            PdfReader reader = new PdfReader(existingPdfPath);
            String tempOutputPath = outputPdfPath + "_temp.pdf";

            // Write to a temporary file first
            try (FileOutputStream tempOutputStream = new FileOutputStream(tempOutputPath)) {
                PdfStamper stamper = new PdfStamper(reader, tempOutputStream);

                PdfContentByte canvas = stamper.getOverContent(1); // Modify page 1
                Image image = Image.getInstance(generatedImagePath);
                image.scaleToFit(500, 300); // Scale image
                image.setAbsolutePosition(xPosition, yPosition); // Position it
                canvas.addImage(image);

                stamper.close(); // Save changes to temp file
            }
            reader.close();

            // Rename temp file to final output
            File tempFile = new File(tempOutputPath);
            File outputFile = new File(outputPdfPath);



            System.out.println("Image added successfully to the PDF at: " + outputPdfPath);
            return tempOutputPath; // Return the final path

        } catch (Exception e) {
            e.printStackTrace();
            return null; // In case of error
        }
    }
    public static String addImageToExistingPdf2(String existingPdfPath, String outputPdfPath, String generatedImagePath, float xPosition, float yPosition) {
        try {
            // Use PdfReader and PdfStamper for modification
            PdfReader reader = new PdfReader(existingPdfPath);
            String tempOutputPath = outputPdfPath + "_SNVN Ltd.pdf";

            // Write to a temporary file first
            try (FileOutputStream tempOutputStream = new FileOutputStream(tempOutputPath)) {
                PdfStamper stamper = new PdfStamper(reader, tempOutputStream);

                PdfContentByte canvas = stamper.getOverContent(1); // Modify page 1
                Image image = Image.getInstance(generatedImagePath);
                image.scaleToFit(500, 300); // Scale image
                image.setAbsolutePosition(xPosition, yPosition); // Position it
                canvas.addImage(image);

                stamper.close(); // Save changes to temp file
            }
            reader.close();

            // Rename temp file to final output
            File tempFile = new File(tempOutputPath);
            File outputFile = new File(outputPdfPath);



            System.out.println("Image added successfully to the PDF at: " + outputPdfPath);
            return outputPdfPath; // Return the final path

        } catch (Exception e) {
            e.printStackTrace();
            return null; // In case of error
        }
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

}
