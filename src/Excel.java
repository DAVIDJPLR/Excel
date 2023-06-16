import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class Excel {
    private Excel(){}

    public class csv {
        private csv(){ }

        public static Queue<ArrayList<String>> read(String fileName){
            ArrayList<String> rows = new ArrayList<>();

            try{
                File in = new File(fileName);
                Scanner myScan = new Scanner(in);
                while (myScan.hasNextLine()){
                    rows.add(myScan.nextLine());
                }
                myScan.close();
            } catch (FileNotFoundException e){
                System.out.println(e.getMessage());
            }

            Queue<ArrayList<String>> out = new LinkedList<>();

            for (String x : rows){
                out.offer(new ArrayList<String>(Arrays.asList(x.split(","))));
            }

            return out;
        }

        public static boolean write(String fileName, Collection<Collection<String>> data){
            Queue<String> rows = new LinkedList<>();
            for (Collection x : data){
                String row = String.join(",", x);
                rows.offer(row);
            }

            try (PrintWriter writer = new PrintWriter(fileName)){
                for (String y : rows){
                    writer.write(y);
                }

                return true;
            } catch (FileNotFoundException e){
                System.out.println(e.getMessage());

                return false;
            }
        }

        public static boolean write(String fileName, Collection<String> data, String delimeter){
            return Excel.csv.write(fileName, convert(data, delimeter));
        }
    }

    public class xls {
        private xls(){}

        public static Queue<ArrayList<String>> read(String fileName){


            Queue<ArrayList<String>> out = new LinkedList<>();

            try{
                FileInputStream file = new FileInputStream(fileName);

                HSSFWorkbook workbook = new HSSFWorkbook(file);

                HSSFSheet sheet = workbook.getSheetAt(0);

                Iterator<Row> rowIterator = sheet.iterator();
                while(rowIterator.hasNext()){
                    Row row = rowIterator.next();
                    ArrayList<String> dataRow = new ArrayList<>();

                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()){
                        Cell cell = cellIterator.next();


                        switch (cell.getCellType())
                        {
                            case NUMERIC:
                                dataRow.add(String.valueOf(cell.getNumericCellValue()));
                                break;
                            case STRING:
                                dataRow.add(cell.getStringCellValue());
                                break;
                        }
                    }
                    out.offer(dataRow);
                }
                file.close();
                return out;
            } catch (FileNotFoundException e) {
                System.out.println(e.getMessage());
                return null;
            } catch (IOException e) {
                System.out.println(e.getMessage());
                return null;
            }
        }

        public static boolean write(String fileName, Collection<Collection<String>> data){

            HSSFWorkbook workbook = new HSSFWorkbook();

            HSSFSheet sheet = workbook.createSheet((fileName.split("\\."))[0]);

            int rownum = 0;
            for (Collection dataRow : data){
                Row row = sheet.createRow(rownum++);
                int cellnum = 0;
                for (Object str : dataRow){
                    Cell cell = row.createCell(cellnum++);
                    cell.setCellValue((String)str);
                }
            }
            try{
                FileOutputStream out = new FileOutputStream(fileName);
                workbook.write(out);
                out.close();
                return true;
            } catch (FileNotFoundException e) {
                System.out.println(e.getMessage());
                return false;
            } catch (IOException e) {
                System.out.println(e.getMessage());
                return false;
            }
        }

        public static boolean write(String fileName, Collection<String> data, String delimeter){
            return Excel.xls.write(fileName, convert(data, delimeter));
        }
    }

    public class xlsx {
        private xlsx(){}

        public static Queue<ArrayList<String>> read(String fileName){

            Queue<ArrayList<String>> out = new LinkedList<>();

            try{
                FileInputStream file = new FileInputStream(fileName);

                XSSFWorkbook workbook = new XSSFWorkbook(file);

                XSSFSheet sheet = workbook.getSheetAt(0);

                Iterator<Row> rowIterator = sheet.iterator();
                while(rowIterator.hasNext()){
                    Row row = rowIterator.next();
                    ArrayList<String> dataRow = new ArrayList<>();

                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()){
                        Cell cell = cellIterator.next();


                        switch (cell.getCellType())
                        {
                            case NUMERIC:
                                dataRow.add(String.valueOf(cell.getNumericCellValue()));
                                break;
                            case STRING:
                                dataRow.add(cell.getStringCellValue());
                                break;
                        }
                    }
                    out.offer(dataRow);
                }
                file.close();
                return out;
            } catch (FileNotFoundException e) {
                System.out.println(e.getMessage());
                return null;
            } catch (IOException e) {
                System.out.println(e.getMessage());
                return null;
            }
        }

        public static boolean write(String fileName, Collection<Collection<String>> data){

            SXSSFWorkbook workbook = new SXSSFWorkbook();

            SXSSFSheet sheet = workbook.createSheet((fileName.split("\\."))[0]);

            int rownum = 0;
            for (Collection dataRow : data){
                Row row = sheet.createRow(rownum++);
                int cellnum = 0;
                for (Object str : dataRow){
                    Cell cell = row.createCell(cellnum++);
                    cell.setCellValue((String)str);
                }
            }
            try{
                FileOutputStream out = new FileOutputStream(fileName);
                workbook.write(out);
                out.close();
                return true;
            } catch (FileNotFoundException e) {
                System.out.println(e.getMessage());
                return false;
            } catch (IOException e) {
                System.out.println(e.getMessage());
                return false;
            }
        }

        public static boolean write(String fileName, Collection<String> data, String delimeter) {
            return Excel.xlsx.write(fileName, convert(data, delimeter));
        }
    }

    private static Collection<Collection<String>> convert(Collection<String> data, String delimeter){

        ArrayList<Collection<String>> out = new ArrayList<>();

        for (String x : data){
            List<String> row =  Arrays.asList(x.split(delimeter));
            out.add(row);
        }

        return out;
    }

    public static Queue<ArrayList<String>> read(String fileName){

        String fileExtension = fileName.split(".")[1];

        if (fileExtension.equalsIgnoreCase("CSV")){
            return Excel.csv.read(fileName);
        } else if (fileExtension.equalsIgnoreCase("XLS")){
            return Excel.xls.read(fileName);
        } else if (fileExtension.equalsIgnoreCase("XLSX")){
            return Excel.xlsx.read(fileName);
        } else {
            return null;
        }
    }

    public static boolean write(String fileName, Collection<Collection<String>> data){

        String fileExtension = fileName.split(".")[1];

        if (fileExtension.equalsIgnoreCase("CSV")){
            return Excel.csv.write(fileName, data);
        } else if (fileExtension.equalsIgnoreCase("XLS")){
            return Excel.xls.write(fileName, data);
        } else if (fileExtension.equalsIgnoreCase("XLSX")){
            return Excel.xlsx.write(fileName, data);
        } else {
            return false;
        }
    }

    public static boolean write(String fileName, Collection<String> data, String delimeter){

        String fileExtension = fileName.split(".")[1];

        if (fileExtension.equalsIgnoreCase("CSV")){
            return Excel.csv.write(fileName, data, delimeter);
        } else if (fileExtension.equalsIgnoreCase("XLS")){
            return Excel.xls.write(fileName, data, delimeter);
        } else if (fileExtension.equalsIgnoreCase("XLSX")){
            return Excel.xlsx.write(fileName, data, delimeter);
        } else {
            return false;
        }
    }
}
