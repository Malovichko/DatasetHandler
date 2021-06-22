import java.io.*;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class FileWorker {
    FileInputStream file;
    HSSFSheet sheet;
    HSSFWorkbook workbook;

    public FileWorker(){
        try {
            openFile();
            workerHSSF();
            dataHSSFWorker(workbook);
            closeFile(workbook);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public void openFile() throws FileNotFoundException {
        file = new FileInputStream("dataset.xls");
    }

    public void closeFile(HSSFWorkbook workbook) throws IOException {
        file.close();
        FileOutputStream outFile = new FileOutputStream("outputDataset.xls");
        workbook.write(outFile);
        outFile.close();
    }

    public void workerHSSF() throws IOException {
        workbook = new HSSFWorkbook(file);
        sheet = workbook.getSheet("VU");
    }

    public void dataHSSFWorker(HSSFWorkbook workbook){
        sheet = workbook.getSheet("VU");

        InputDatasetHandler inputDatasetHandler = new InputDatasetHandler();

        inputDatasetHandler.deleteWithoutShift(sheet);
        inputDatasetHandler.shiftRows(sheet);
        inputDatasetHandler.calcValuesKGF(sheet);
        inputDatasetHandler.deleteLastColumn(sheet);
        inputDatasetHandler.toNull(sheet);
        inputDatasetHandler.printColumn(sheet);
        inputDatasetHandler.deleteUnnecessaryColumns(sheet);
    }
}
