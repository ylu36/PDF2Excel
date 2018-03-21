import java.util.*;
import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
class Benefix {
    static String text;
    static int rowNo;
    static String body[] = new String[5];
    static Double data[] = new Double[47];

    public static void parseEachPage() {
        int index = text.indexOf("Valid for Effective Dates");
        body[0] = text.substring(index+27, index+37);
        body[1] = text.substring(index+41, index+51);
        index = text.indexOf("Plan Name");
        body[2] = text.substring(index+10, index+41);
        body[3] = "PA";
        index = text.indexOf("PARA");
        body[4] = text.substring(index+4, index+6);
        //System.out.println(text);

        index = text.indexOf("0-20");
        text = text.substring(index+5);
        text = text.replaceAll("\\+"," ");
       // System.out.println(text);
        Scanner s = new Scanner(text);
        for(int i = 1; i < 46; i ++) {
            data[i] = s.nextDouble();
            // System.out.println(i+" "+data[i]);
            if(s.hasNextInt())
                s.nextInt();
            else break;
        }
        data[0] = data[1];
        data[46] = data[45];
        // index = text.indexOf("area definitions");
        //text = text.substring(index+10);
    }

    public static void writeXLSXFile(int index) throws IOException {
        FileInputStream fsIP = new FileInputStream(new File("BeneFix Small Group Plans upload template.xlsx"));
        XSSFWorkbook wb = new XSSFWorkbook(fsIP);
        XSSFSheet sheet = wb.getSheetAt(0);
        Row row = null;
        Cell cell = null;
        row = sheet.getRow(index);
        if (row == null) {
            row = sheet.createRow(index);
        }
        for (int i = 0; i < 5; i++) {
            cell = row.getCell(i);
            if (cell == null) {
                cell = row.createCell(i);
            }
            cell.setCellValue(body[i]);
        }
        for(int i = 0; i < 47; i++) {
            cell = row.getCell(i+5);
            if (cell == null) {
                cell = row.createCell(i+5);
            }
            cell.setCellValue(data[i]);
        }
        fsIP.close();
        FileOutputStream output_file =new FileOutputStream(new File("BeneFix Small Group Plans upload template.xlsx"));
        wb.write(output_file);
        output_file.close();
    }

    public static void main(String[] s) throws IOException {
        rowNo = 0;
        for (int paraNo = 1; paraNo <= 9; paraNo++) {
            if(paraNo == 4) continue;
            String fileName = "para0" + paraNo + ".pdf";
            PDDocument document = PDDocument.load(new File(fileName));
            if (!document.isEncrypted()) {
                System.out.println("Now reading " + fileName + "...");
                PDFTextStripper stripper = new PDFTextStripper();
                text = stripper.getText(document);
                // System.out.println(text.length());
            }
            for (int i = 1; i <= 45; i++) {
                parseEachPage();
                System.out.println("Now writing " + fileName + " page " + i + " to Excel...");
                writeXLSXFile(i+rowNo);
            }
            rowNo += 45;
            // System.out.println(rowNo);
            document.close();
        }
        System.out.println("=======================\n" + "pdf to excel finished...");
    }
}