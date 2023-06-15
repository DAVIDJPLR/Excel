import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.PrintWriter;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

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
            Queue<String> rows = new LinkedList<>();
            for (String x : data){
                String[] rowList = x.split(delimeter);
                String row = String.join(",", rowList);
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
    }

    public class xls {
        private xls(){}

        public static Queue<ArrayList<String>> read(String fileName){
            /**
             * ToDo: Code
             */
            try {
                HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(fileName));
                HSSFSheet myExcelSheet = myExcelBook.getSheet("Birthdays");
                HSSFRow row = myExcelSheet.getRow(0);

                if (row.getCell(0).getCellType() == HSSFCell.CELL_TYPE_STRING) {
                    String name = row.getCell(0).getStringCellValue();
                    System.out.println("name : " + name);
                }

                if (row.getCell(1).getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                    Date birthdate = row.getCell(1).getDateCellValue();
                    System.out.println("birthdate :" + birthdate);
                }

                myExcelBook.close();
            } catch (FileNotFoundException e){
                System.out.println(e.getMessage());
            }

            return null;
        }

        public static boolean write(String fileName, Collection<Collection<String>> data){
            /**
             * ToDo: Code
             */
            return false;
        }

        public static boolean write(String fileName, Collection<String> data, String delimeter){
            /**
             * ToDo: Code
             */
            return false;
        }
    }

    public class xlsx {
        private xlsx(){}

        public static Queue<ArrayList<String>> read(String fileName){
            /**
             * ToDo: Code
             */
            return null;
        }

        public static boolean write(String fileName, Collection<Collection<String>> data){
            /**
             * ToDo: Code
             */
            return false;
        }

        public static boolean write(String fileName, Collection<String> data, String delimeter){
            /**
             * ToDo: Code
             */
            return false;
        }
    }

    private Collection<Collection<String>> convert(Collection<String> data, String delimeter){
        /**
         * ToDo: Code
         */
        return null;
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
