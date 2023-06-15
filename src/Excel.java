import java.util.ArrayList;
import java.util.Collection;
import java.util.Queue;

public class Excel {
    private Excel(){}

    public class csv {
        private csv(){ }

        public static Collection<Collection<String>> read(String fileName){
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

    public class xls {
        private xls(){}

        public static Collection<Collection<String>> read(String fileName){
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

        public static Collection<Collection<String>> read(String fileName){
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

    public Collection<Collection<String>> read(String fileName){

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

    public boolean write(String fileName, Collection<Collection<String>> data){

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

    public boolean write(String fileName, Collection<String> data, String delimeter){

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
