import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;

public class DocxToPdf {
    public static void main(String[] args) throws IOException {
        File input = new File("./input");
        if (!input.isDirectory()) {
            System.out.println("can't find input directory");
            return;
        }
        File output = new File("./output");
        if (!output.isDirectory()) {
            if (!output.mkdirs()) {
                System.out.println("can't create output directory");
                return;
            }
        }

        PDFMergerUtility PDFmerger = new PDFMergerUtility();
        PDFmerger.setDestinationFileName(output.getAbsolutePath() + "/merge.txt");

        File[] files = input.listFiles();
        for (File f : files) {
            File pdf = new File(output, f.getName() + ".pdf");
            ConvertToPDF(f, pdf);
            PDFmerger.addSource(pdf);
        }
        PDFmerger.mergeDocuments();
    }

    public static void ConvertToPDF(File input, File output) {
        try {
            InputStream doc = new FileInputStream(input);
            XWPFDocument document = new XWPFDocument(doc);
            PdfOptions options = PdfOptions.create();
            OutputStream out = new FileOutputStream(output);
            PdfConverter.getInstance().convert(document, out, options);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
