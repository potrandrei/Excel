package Excel;


import jxl.Workbook;
import jxl.write.*;
import jxl.write.Boolean;


import java.io.File;
import java.lang.Number;
import java.util.Date;

public class App {
        public static void main( String[] args ) throws Exception {
        File f = new File("D:\\Programare\\Programe\\IntelliJ\\Excel\\excel.xls");

                WritableWorkbook myexcel = Workbook.createWorkbook(f);
                WritableSheet mysheet = myexcel.createSheet("mysheet",0);
                Label l = new Label(0,0,"data 1");
                mysheet.addCell(new Label(0, 0, "ABC"));
                mysheet.addCell(new Label(1, 0, "DEF"));

                mysheet.addCell(new Boolean(0, 1, true));
                mysheet.addCell(new Boolean(1, 1, false));

                mysheet.addCell(new DateTime(0, 2, new Date()));
                mysheet.addCell(new DateTime(1, 2, new Date(System.currentTimeMillis() + 90000000)));

                myexcel.write();
                myexcel.close();

                System.out.println("The file was created and updated.");


        }
}
