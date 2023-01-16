package ReadExcel;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
public class Read_Excel {
    public static void main(String[] args) throws IOException {
        File file = new File("./Data/2023 Monthly Link Targets.xlsx");
        FileInputStream Gowtham = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(Gowtham);
        Sheet sheet = workbook.getSheetAt(0);
        Map<String, String> data = new HashMap<>();
        int rows = sheet.getLastRowNum();
        for (Row row : sheet) {
            for (Cell cell : row) {
                String columnName = sheet.getRow(0).getCell(cell.getColumnIndex()).getStringCellValue();
                String cellValue = "";
                switch (cell.getCellType()) {
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue()+"  ");
                        break;
                    case STRING:
                    	System.out.print(cell.getStringCellValue()+"  ");
                        break;
                }
            }         
            System.out.println();
        }
        Gowtham.close();
    }
}

