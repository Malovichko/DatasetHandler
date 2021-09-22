import java.io.*;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.json.simple.parser.ParseException;

public class FileWorker {
    FileInputStream file;
    HSSFSheet sheet;
    HSSFWorkbook workbook;

    public FileWorker(){
        try {
            openFile();
            workerHSSF();
            try {
                dataHSSFWorker(workbook);
            } catch (ParseException e) {
                e.printStackTrace();
            }
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
        sheet = workbook.getSheet("Лист 1");
    }

    public void dataHSSFWorker(HSSFWorkbook workbook) throws IOException, ParseException {
        sheet = workbook.getSheet("Лист 1");
        InputDatasetHandler inputDatasetHandler = new InputDatasetHandler();
        inputDatasetHandler.getRow(sheet);
    }
}
