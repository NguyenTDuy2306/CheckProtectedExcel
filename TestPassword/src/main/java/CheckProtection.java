import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;
import java.util.Iterator;
import java.util.Scanner;

public class CheckProtection {
    String path;

    public boolean isEncrypted(String path) {
        this.path = path;
        try {
            try {
                new POIFSFileSystem(new FileInputStream(path));
            } catch (IOException ex) {

            }
            System.out.println("this file has protected by password");
            return true;
        } catch (OfficeXmlFileException e) {
            System.out.println("this file not protected");
            return false;
        }
    }

    public CheckProtection() {
    }

    public void inputPassword(String path) {
        Scanner sc = new Scanner(System.in);
        String password = sc.nextLine();
        File xlsxFile = new File(path);
        try {
            POIFSFileSystem fs = new POIFSFileSystem(xlsxFile);
            EncryptionInfo info = new EncryptionInfo(fs);
            Decryptor decryptor = Decryptor.getInstance(info);

            //Verifying the password
            if (!decryptor.verifyPassword(password)) {
                throw new RuntimeException("Incorrect password: Unable to process");
            }

            InputStream dataStream = decryptor.getDataStream(fs);

            XSSFWorkbook workbook = new XSSFWorkbook(dataStream);
            Sheet sheet = workbook.getSheetAt(0);

            Iterator<Row> iterator = sheet.iterator();

            while (iterator.hasNext()) {
                Row nextRow = iterator.next();
                Iterator<Cell> cellIterator = nextRow.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue());
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue());
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue());
                            break;
                        default:
                            break;
                    }
                    System.out.print(" ");
                }
                System.out.println();
            }
            workbook.close();
        } catch (GeneralSecurityException | IOException ex) {
            throw new RuntimeException("Unable to process encrypted document", ex);
        }
    }


}
