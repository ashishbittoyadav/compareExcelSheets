import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;




public class Main {

    static ArrayList<DataObject> Excel1 = new ArrayList<DataObject>();

    static ArrayList<DataObject> Excel2 = new ArrayList<DataObject>();

    static ArrayList<DataObject> Excel3 = new ArrayList<DataObject>();

    static int numberOfCells = 0;

    static int numberOfRows = 0;

    int numberOfRows2 = 0;

    int numberOfCells2 = 0;

    static int total = 0;

    static int error = 0;

    static int errorpercent = 0;

    static String ResultString1 = null;

    static String ResultString2 = null;

    static String Filename1 = null;

    static String Filename2 = null;

    static String Filename3 = null;

    public static void main(String[] args) throws IOException {

        String ExcelPath1="D:\\new\\Book1.xlsx";

        String ExcelPath2="D:\\new\\Book2.xlsx";

        String ExcelPath3="D:\\new\\Book3.xlsx";

        Scanner sc = new Scanner(System.in);

//        System.out.println("Enter the Path of first file:- \n Eg:- C:\\Users\\ruturaj.tambe\\Desktop\\Excel1.xlsx");
//
//        String Filename1 = sc.nextLine();
//
//        System.out.println("Enter the Path of second file:- ");
//
//        String Filename2 = sc.nextLine();
//
//        System.out.println("Enter the name of Excel sheet to be generated:-");
//
//        String Filename3 = sc.nextLine();

        new Main().readExcel(ExcelPath1, ExcelPath2, ExcelPath3);

        // new ReadExcelFinal().print(Excel1);

        // new ReadExcelFinal().print(Excel2);

        sc.close();

    }

    public void writeExcel(String Filename3) {

        try {

            XSSFWorkbook workbook3 = new XSSFWorkbook();

            XSSFSheet sheet3 = workbook3.createSheet("Sheet3");

            int q = 0;

            int rows = 0;

            System.out.println("\n\nCreating new excel sheet:");

            for (int i = 0; i <= numberOfRows; i++) {

                XSSFRow row = sheet3.createRow(rows++);

                for (int j = 0; j < numberOfCells; j++) {

                    XSSFCell cell = row.createCell(j);

                    cell.setCellValue(Excel3.get(q).getValue());

                    q++;

                }

            }

            FileOutputStream out = new FileOutputStream(new File(Filename3 + ".xlsx"));

            workbook3.write(out);

            out.close();

//            workbook3.close();

            new Main().printString(Excel3);

        } catch (Exception e) {

            // TODO Auto-generated catch block

            e.printStackTrace();

        }

    }

    public void compareStore(ArrayList<DataObject> array1, ArrayList<DataObject> array2, String Filename3) throws IOException {

        try {

            int r1 = array1.size();

            System.out.println("Total numbers Values in first excel:" + r1);

            int r2 = array2.size();

            System.out.println("Total numbers Values in second excel:" + r2);

            int value = 0;


            Excel3.addAll(array1);
            Excel3.addAll(array2);

            System.out.println(Excel3.removeAll(array2));

//            for (int i=0;i<r2;i++) {
//                for (int j = 0; j < Excel3.size(); j++) {
//                    if (Excel3.get(j).getValue().equals(array2.get(i).getValue())) {
//                        System.out.println(array2.get(i).getValue());
////                        Excel3.add(array2.get(i).toString());
////                        value++;
////                        total++;
//                        break;
//                    }
//                    else{
////                        System.out.println("else "+array2.get(i).getValue());
//                        Excel3.add(array2.get(i));
////                        value++;
////                        total++;
//                        break;
//                    }
//                }
//            }


            System.out.println("\nNo of cells that did not match:-  " + error);

            System.out.println("Percent error in the sheets:-  " + (error * 100.00 / total) + " %");

//            new Main().writeExcel(Filename3);

        } catch (Exception e) {

            System.out.println(e);

        }

        System.out.println(Excel3.size());

//        for (DataObject data :
//                Excel3) {
//            System.out.println(data.getValue());
//        }
    }

    public void print(ArrayList<Object> array) {

        try {

            int r1 = array.size();

            System.out.println("Total Number of elements:" + r1);

            int s = 0;

            System.out.println("Printing the contents of excel sheet:");

            for (int i = 0; i <= numberOfRows; i++) {

                System.out.println("");

                for (int j = 1; j <= numberOfCells; j++) {

                    System.out.print(array.get(s) + "\t\t\t");

                    s++;

                }

            }

        } catch (Exception e) {

            e.printStackTrace();

        }

    }

    public void printString(ArrayList<DataObject> array) {

        try {

            int r1 = array.size();

            System.out.println("Size Of first array:" + r1);

            int s = 0;

            System.out.println("Printing the contents of excel sheet:");

            for (int i = 0; i <= numberOfRows; i++) {

                System.out.println("");

                for (int j = 1; j <= numberOfCells; j++) {

                    System.out.print(array.get(s) + "\t\t\t");

                    s++;

                }

            }

        } catch (Exception e) {

            e.printStackTrace();

        }

    }

    public void readExcel(String Filename1, String Filename2, String Filename3) {

        try {

            FileInputStream file1 = new FileInputStream(new File(Filename1));

            FileInputStream file2 = new FileInputStream(new File(Filename2));

            System.out.println(Filename1);

            System.out.println(Filename2);

            Pattern regex1 = Pattern.compile("([^\\\\/:*?\"<>|\r\n]+$)");

            Matcher regexMatcher1 = regex1.matcher(Filename1);

            if (regexMatcher1.find()) {

                ResultString1 = regexMatcher1.group(1);

                System.out.println(ResultString1);

            }

            Pattern regex2 = Pattern.compile("([^\\\\/:*?\"<>|\r\n]+$)");

            Matcher regexMatcher2 = regex2.matcher(Filename2);

            if (regexMatcher2.find()) {

                ResultString2 = regexMatcher2.group(1);

                System.out.println(ResultString2);

            }

            //Create Workbook instance holding reference to .xlsx file

            final XSSFWorkbook workbook1 = new XSSFWorkbook(file1);

            final XSSFWorkbook workbook2 = new XSSFWorkbook(file2);

            //Get first/desired sheet from the workbook

            final XSSFSheet sheet1 = workbook1.getSheetAt(0);

            final XSSFSheet sheet2 = workbook2.getSheetAt(0);

            //Iterate through each rows one by one

            final Iterator<org.apache.poi.ss.usermodel.Row> rowIterator1 = sheet1.iterator();

            final Iterator<org.apache.poi.ss.usermodel.Row> rowIterator2 = sheet2.iterator();

            final Iterator<org.apache.poi.ss.usermodel.Row> rowIterator1_1 = sheet1.iterator();

            final Iterator<org.apache.poi.ss.usermodel.Row> rowIterator2_1 = sheet2.iterator();

            numberOfRows = sheet1.getLastRowNum();

            if (rowIterator1_1.hasNext()) {

                org.apache.poi.ss.usermodel.Row headerRow1 = rowIterator1_1.next();   //get the number of cells in the header row

                numberOfCells = headerRow1.getPhysicalNumberOfCells();

            }

            System.out.println("Number of rows :" + numberOfRows);

            System.out.println("Number of cells :" + numberOfCells);

            numberOfRows2 = sheet2.getLastRowNum();

            if (rowIterator2_1.hasNext()) {

                org.apache.poi.ss.usermodel.Row headerRow2 = rowIterator2_1.next();   //get the number of cells in the header row

                numberOfCells2 = headerRow2.getPhysicalNumberOfCells();

            }

            System.out.println("Number of rows :" + numberOfRows2);

            System.out.println("Number of cells :" + numberOfCells2);

            if (numberOfRows == numberOfRows2 && numberOfCells == numberOfCells2) {

                while (rowIterator1.hasNext() && rowIterator2.hasNext()) {

                    org.apache.poi.ss.usermodel.Row row1 = rowIterator1.next();

                    org.apache.poi.ss.usermodel.Row row2 = rowIterator2.next();

                    //For each row, iterate through all the columns

                    Iterator<Cell> cellIterator1 = row1.cellIterator();

                    Iterator<org.apache.poi.ss.usermodel.Cell> cellIterator2 = row2.cellIterator();

                    while (cellIterator1.hasNext() && cellIterator2.hasNext()) {

                        org.apache.poi.ss.usermodel.Cell cell1 = cellIterator1.next();

                        org.apache.poi.ss.usermodel.Cell cell2 = cellIterator2.next();

                        //Check the cell type and format accordingly

                        DataObject dataObject = new DataObject();
                        switch (cell1.getCellTypeEnum()) {


                            case NUMERIC:

                                //  System.out.print(cell1.getNumericCellValue() + "\t\t");
                                dataObject.setIndex(cell1.getNumericCellValue());
//                                Excel1.add(cell1.getNumericCellValue());

                                break;

                            case STRING:

                                //  System.out.print(cell1.getStringCellValue() + "\t\t");

                                dataObject.setValue(cell1.getStringCellValue());
//                                Excel1.add(cell1.getStringCellValue());

                                break;
                        }
                        Excel1.add(dataObject);

                        switch (cell2.getCellTypeEnum()) {

                            case NUMERIC:

                                //  System.out.print(cell2.getNumericCellValue() + "\t\t");
                                dataObject.setIndex(cell2.getNumericCellValue());
//                                Excel2.add(cell2.getNumericCellValue());

                                break;

                            case STRING:

                                //  System.out.print(cell2.getStringCellValue() + "\t\t");

                                dataObject.setValue(cell2.getStringCellValue());
//                                Excel2.add(cell2.getStringCellValue());

                                break;

                        }
                        Excel2.add(dataObject);
                    }

                    //System.out.println("");

                }

                System.out.println("\nRead Complete: Values from ExcelSheet 1 are stored in Excel1 and Values from ExcelSheet 2 are stored in Excel2 \n");

                file1.close();

                file2.close();

//                workbook1.close();

//                workbook2.close();

//                new Main().compareStore(Excel1, Excel2, Filename3);

            } else {

                System.out.println("Rows and Columns do not match");

                while (rowIterator2.hasNext()) {

                    org.apache.poi.ss.usermodel.Row row2 = rowIterator2.next();
                    Iterator<org.apache.poi.ss.usermodel.Cell> cellIterator2 = row2.cellIterator();

                    while (cellIterator2.hasNext()){
                        org.apache.poi.ss.usermodel.Cell cell2 = cellIterator2.next();
                        DataObject dataObject = new DataObject();
                        switch (cell2.getCellTypeEnum()) {

                            case NUMERIC:

                                //  System.out.print(cell2.getNumericCellValue() + "\t\t");

                                dataObject.setIndex(cell2.getNumericCellValue());
//                                Excel2.add(cell2.getNumericCellValue());

                                break;

                            case STRING:

                                //  System.out.print(cell2.getStringCellValue() + "\t\t");

                                dataObject.setValue(cell2.getStringCellValue());
//                                Excel2.add(cell2.getStringCellValue());

                                break;

                        }
                        Excel2.add(dataObject);
                    }
                }

                while (rowIterator1.hasNext()) {
                    org.apache.poi.ss.usermodel.Row row1 = rowIterator1.next();
                    Iterator<Cell> cellIterator1 = row1.cellIterator();
                    //For each row, iterate through all the columns
                    while (cellIterator1.hasNext()) {
                        org.apache.poi.ss.usermodel.Cell cell1 = cellIterator1.next();

                        //Check the cell type and format accordingly

                        DataObject dataObject = new DataObject();
                        switch (cell1.getCellTypeEnum()) {

                            case NUMERIC:

                                //  System.out.print(cell1.getNumericCellValue() + "\t\t");

                                dataObject.setIndex(cell1.getNumericCellValue());
//                                Excel1.add(cell1.getNumericCellValue());

                                break;

                            case STRING:

                                //  System.out.print(cell1.getStringCellValue() + "\t\t");

                                dataObject.setValue(cell1.getStringCellValue());
//                                Excel1.add(cell1.getStringCellValue());

                                break;

                        }
                        Excel1.add(dataObject);
                    }

                    new Main().compareStore(Excel1,Excel2,Filename3);
                    //System.out.println("");

                }

//                for (Object data:
//                     Excel1) {
//                    System.out.println(data.toString());
//                }
//                for (Object data:
//                        Excel2) {
//                    System.out.println(data.toString());
//                }
            }

        } catch (Exception e) {

            e.printStackTrace();

        }

    }

}

class DataObject{
    private double index;
    private String value;

    void setIndex(double _index){
        index = _index;
    }

    void setValue(String _value){
        value = _value;
    }

    String getValue(){
        return value;
    }

    double getIndex(){
        return index;
    }
}
