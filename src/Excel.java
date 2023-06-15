import java.io.File;
import java.io.FileNotFoundException;
import java.io.PrintWriter;
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
