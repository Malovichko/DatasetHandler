import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.Iterator;

public class InputDatasetHandler {
    private ArrayList<Row> findRowsToDelete(HSSFSheet sheet){
        ArrayList<Row> arrayList = new ArrayList<>();
        Iterator<Row> rowIterator = sheet.rowIterator();
        int number = 0;
        for (int i = 0; i < 3; i++) rowIterator.next();

        while (rowIterator.hasNext()){
            Row row = rowIterator.next();
            if ((row.getCell(31).getNumericCellValue()==0.0)&&(row.getCell(32).getNumericCellValue()==0.0)&&
                    (row.getCell(33)==null)) {
                arrayList.add(row);
            }
            if (number == 184) break;
            number++;
        }
        return arrayList;
    }

    private void deleteRows(HSSFSheet sheet, ArrayList<Row> rowsToDelete){
        for (Row row : rowsToDelete)
            sheet.removeRow(row);
    }

    private void printRowCount(HSSFSheet sheet){
        Iterator<Row> iterator = sheet.rowIterator();
        while (iterator.hasNext()){
            iterator.next();
        }
    }
    private ArrayList<Integer> createNotEmptyRowsList(HSSFSheet sheet){
        ArrayList<Integer> list = new ArrayList<>();
        Iterator<Row> iterator = sheet.rowIterator();
        while (iterator.hasNext()){
            Row row = iterator.next();
            list.add(row.getRowNum());
        }
        return list;
    }
    public void shiftRows(HSSFSheet sheet){
        ArrayList<Integer> list = createNotEmptyRowsList(sheet);
        for(int i = sheet.getLastRowNum(); i > 0; i--){
            if(!list.contains(i)){
                sheet.shiftRows(i + 1, sheet.getLastRowNum(), -1);
            }
        }
    }
    public void deleteWithoutShift(HSSFSheet sheet) {
        ArrayList<Row> rows = findRowsToDelete(sheet);
        printRowCount(sheet);
        deleteRows(sheet, rows);
        printRowCount(sheet);
    }
    public void calcValuesKGF(HSSFSheet sheet){
        Iterator<Row> iterator = sheet.rowIterator();
        for (int i = 0; i < 3; i++) iterator.next();
        int number = 3;
        while (iterator.hasNext()){
            Row row = iterator.next();
            if ((row.getCell(32).getNumericCellValue() == 0.0) &&
                    (row.getCell(33) != null)) {
                HSSFCell cell = (HSSFCell) row.getCell(32);
                cell.setCellValue(row.getCell(33).getNumericCellValue() * 1000);
            }
            number++;
            if (number >= 96) break;

        }
    }
    public void deleteLastColumn(HSSFSheet sheet) {
        Iterator<Row> iterator = sheet.rowIterator();
        while (iterator.hasNext()){
            Row row = iterator.next();
            if (row.getCell(33) != null)
                row.removeCell(row.getCell(33));
        }
    }

    public void printColumn(HSSFSheet sheet) {
        Iterator<Row> iterator = sheet.rowIterator();
        int number = 31, i = 0;
        while (iterator.hasNext()){
            i++;
            Row row = iterator.next();
            if ((i >= 4) && (row.getCell(number) != null) && (
                    ((row.getCell(number).getCellType().toString().equals("STRING")) && ((row.getCell(number).getStringCellValue().equals("-"))||(row.getCell(number).getStringCellValue().equals("не спускался"))))
                            ||(row.getCell(number).getCellType().toString().equals("BLANK")) && (row.getCell(number).getNumericCellValue()==0.0)));
        }
    }

    public void toNull(HSSFSheet sheet) {
        Iterator<Row> iterator = sheet.rowIterator();
        int number = 32, i = 0;
        while (iterator.hasNext()){
            i++;
            Row row = iterator.next();
            if ((i >= 4) && (i < 97))
                for (int column_num = 0; column_num <= number; column_num++) {
                    if ((row.getCell(column_num) != null) && (
                            ((row.getCell(column_num).getCellType().toString().equals("STRING")) && ((row.getCell(column_num).getStringCellValue().equals("-")) || (row.getCell(column_num).getStringCellValue().equals("не спускался"))))
                                    || (row.getCell(column_num).getCellType().toString().equals("BLANK")) && (row.getCell(column_num).getNumericCellValue() == 0.0))) {
                        row.removeCell(row.getCell(column_num));
                    }
                }
        }
    }

    private ArrayList<Integer> findHighPercent(HSSFSheet sheet) {
        Iterator<Row> iterator = sheet.rowIterator();
        iterator.next();
        Row row = iterator.next();
        double nullColumn[] = new double[row.getLastCellNum()];
        iterator.next();
        for (int i = 4; i < 97; i++){
            row = iterator.next();
            for (int j = 2; j < 31; j++)
                if (row.getCell(j) == null)
                    nullColumn[j]++;
        }
        double total = 93.0;
        ArrayList<Integer> arrayList = new ArrayList<>();
        for (int i = 2; i < 31; i++) {
            nullColumn[i] = nullColumn[i] / total * 100.0;
            if (nullColumn[i] > 60.0)
                arrayList.add(i);
        }
        return arrayList;
    }

    public void deleteUnnecessaryColumns(HSSFSheet sheet) {
        ArrayList<Integer> columnList = findHighPercent(sheet);
        Iterator<Row> iterator = sheet.rowIterator();
        while (iterator.hasNext()){
            Row row = iterator.next();
            for (int col : columnList)
                if (row.getCell(col) != null)
                    row.removeCell(row.getCell(col));
        }
    }
}
