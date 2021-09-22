import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.*;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

public class InputDatasetHandler {

    private void createJSONObjectThreads(JSONArray jsonArrayThreads, Row row) {
        List<String> threadsTitle = Arrays.asList("DMC", "Gamma", "Madeira", "Anchor", "PNK im. Kirova", "Dimensions", "Dome", "Bucilla", "J&P Coats", "BELKA", "Olympus", "COSMO", "Yeidam (耶单)", "SEMCO");
        for (int i = 0; i < threadsTitle.size(); i++) {
            int cellNum = i + 4;
            createThreadsObject(jsonArrayThreads, threadsTitle.get(i), String.valueOf(row.getCell(cellNum)));
        }
    }

    public static boolean contains(String pattern, String content) {
        return content.matches(pattern);
    }

    private void createThreadsObject(JSONArray jsonArrayThreads, String threadTitle, String number) {
        if (contains("(.*).(.*)", number)) {
            number = number.split("\\.", 2)[0];
        }
        if (!number.equals("nil")) {
            JSONObject jThread = new JSONObject();
            jThread.put("title", threadTitle);
            jThread.put("number", number);
            jsonArrayThreads.add(jThread);
        }

    }

    private JSONObject createJSONObjectColors(String imageName, String red, String green, String blue, String colorsIcon) {
        JSONObject jo = new JSONObject();

        jo.put("imageName", imageName);
        jo.put("colorsIcon", colorsIcon);
        jo.put("size", "5");
        jo.put("red", red);
        jo.put("green", green);
        jo.put("blue", blue);
        return jo;
    }


    public ArrayList<Row> getRow(HSSFSheet sheet) throws IOException, ParseException {

        prepareData(sheet);
        ArrayList<Row> arrayList = new ArrayList<>();
        Iterator<Row> rowIterator = sheet.rowIterator();
        Map<String, String> colorsInApp = getColorsArray();
        int number = 0;

        JSONArray jsonArray = new JSONArray();
        for (int i = 0; i < 1; i++) rowIterator.next();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            String value = String.valueOf(row.getCell(2));
            if (colorsInApp.containsKey(value)) {
                JSONObject jo = getJsonObject(row, colorsInApp.get(value));
                jsonArray.add(jo);
//                System.out.println(row.getCell(2));
            }
            if (number == 447) break;
            number++;
        }

        createFile(jsonArray);
        return arrayList;
    }

    private void createFile(JSONArray jsonArray) throws IOException {
        FileWriter file = new FileWriter("out.json");
        jsonArray.writeJSONString(file);
        file.flush();
        file.close();
    }

    private JSONObject getJsonObject(Row row, String number) {
        String[] rgb = getRGBSubstring(String.valueOf(row.getCell(1)));
        String red = rgb[0];
        String green = rgb[1];
        String blue = rgb[2];

        JSONObject jo = createJSONObjectColors(String.valueOf(row.getCell(2)), red, green, blue, number);
        JSONArray jsonArrayThreads = new JSONArray();
        createJSONObjectThreads(jsonArrayThreads, row);
        jo.put("thread", jsonArrayThreads);
        return jo;
    }

    private String[] getRGBSubstring(String rgb) {
        String[] subStr;
        String delimiter = ",";
        subStr = rgb.split(delimiter);
        for (int i = 0; i < subStr.length; i++) {
//            System.out.println(subStr[i]);
        }
        return subStr;
    }

    public void prepareData(HSSFSheet sheet) {
        Iterator<Row> rowIterator = sheet.rowIterator();
        int number = 0;
        for (int i = 0; i < 1; i++) rowIterator.next();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            for (int i = 0; i < row.getLastCellNum(); i++) {
                Cell cell = row.getCell(i);
                if (String.valueOf(cell).equals("—") || String.valueOf(cell).equals("") || String.valueOf(cell).equals("\u2014")) {
                    cell.setCellValue("nil");
                }
            }
            if (number == 447) {
                break;
            }
            number++;
        }
    }

    public Map<String, String> getColorsArray() throws IOException, ParseException {
        JSONParser parser = new JSONParser();
        Map<String, String> colors = new HashMap<>();
        JSONArray a = (JSONArray) parser.parse(new FileReader("list_threads.json"));

        for (Object o : a)
        {
            JSONObject thread = (JSONObject) o;

            String name = (String) thread.get("imageName");
            String icon = (String) thread.get("colorsIcon");
//            System.out.println(name);
            colors.put(name, icon);
        }
        return colors;
    }
}
