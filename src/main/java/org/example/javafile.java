package org.example;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.pdf.ColumnText;
import java.io.BufferedReader;
import java.io.FileInputStream;

import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;

public class javafile {

    public static void main(String[] args) {
        String csvFile = "E:\\pdf/certificatedata2.csv";
        String line;
        String csvSplitBy = ",";

        try (BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(csvFile), "UTF-8"))) {
           br.readLine();
            while ((line = br.readLine()) != null) {

                String[] data = line.split(csvSplitBy);
                if (data.length >= 2) {
                    String inputPdf = "E:\\pdf/rcertificate.pdf";
                    String outputPdf = "E:\\pdf/ics/"+data[0]+".pdf";
                    String customFontPath = "E:\\pdf/Anydore.otf";// Update with the correct font file path
                    try {
                        PdfReader reader = new PdfReader(inputPdf);

                        Document document = new Document(PageSize.A4);
                        PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(outputPdf));
                        document.open();

                        PdfContentByte canvas = writer.getDirectContent();

                        document.newPage();
                        PdfImportedPage page = writer.getImportedPage(reader, 1);
                        canvas.addTemplate(page, 0, 0);
                        BaseFont customFont = BaseFont.createFont(customFontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

                        // Add text in the middle of the page
                        int xPosition = (int) ((document.right() + document.left()) / 2);
                        int yPosition = 125 + (int) ((document.top() + document.bottom()) / 2);

                        // Add the first text
                        String text1 = data[1];
                        Font font1 = new Font(customFont, 16);
                        font1.setColor(new BaseColor(0, 0, 128));
                        Paragraph paragraph1 = new Paragraph(text1, font1);
                        paragraph1.setAlignment(Element.ALIGN_JUSTIFIED);
                        paragraph1.setLeading(14);
                        ColumnText.showTextAligned(canvas, Element.ALIGN_CENTER, paragraph1, xPosition+30, yPosition + 40, 0);

                        // Add the second text
                        String text2 = data[0];
                        Font font2 = new Font(customFont, 16);
                        font2.setColor(new BaseColor(0, 0, 128));
                        Paragraph paragraph2 = new Paragraph(text2, font2);
                        paragraph2.setAlignment(Element.ALIGN_JUSTIFIED);
                        paragraph2.setLeading(14);
                        ColumnText.showTextAligned(canvas, Element.ALIGN_CENTER, paragraph2, xPosition + 205, yPosition + 40, 0);

                        document.close();
                        writer.close();
                        reader.close();

                        System.out.println("File Id: "+data[0]);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    
                   
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
