
package laba;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Iterator;
import java.util.Scanner;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;





public class Laba {

  
    public static void main(String[] args) throws Exception {
        int pvt;
        do{                       
        Scanner in = new Scanner(System.in);        
        System.out.print("1. Продажа товара\n2. Поставка товара\n3. Данные по продажам\n ");
        int b = in.nextInt();
        switch (b){
            case 1:
                sale sl = new sale();
                sl.excel (); 
                sl.sklad();
                break;
            case 2:  
                post post = new post();
                post.post();               
                break;
            case 3:
                jur jr= new jur();
                jr.excel();
                break;
        }
        System.out.println("Нажмите 1 чтобы повторить. Нажмите 2 чтобы закончить.");
         pvt=in.nextInt();
        }while(pvt==1);
    
    }
}
class sale{
    public void excel() throws Exception {
        
        InputStream in = new FileInputStream("D:\\1.xls");
        HSSFWorkbook wb = new HSSFWorkbook(in);
 
        Sheet sheet = wb.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        while (it.hasNext()) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            while (cells.hasNext()) {
                Cell cell = cells.next();
                int cellType = cell.getCellType();
                switch (cellType) {
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(cell.getStringCellValue() + "||");
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print("[" + cell.getNumericCellValue() + "]");
                        break;
    
                }
            }
            System.out.println();
        }
    }
    public void sklad() throws Exception {
        Scanner in = new Scanner(System.in);       
        InputStream inp = new FileInputStream("D:\\1.xls");
        Workbook book = new HSSFWorkbook(inp);        
        System.out.print("Что продать?Введите ID \n ");
        int pr1 = in.nextInt();
        System.out.print("Сколько\n ");
        int pr2 = in.nextInt();
        Sheet sheet = book.getSheetAt(0);
        Row row = sheet.getRow(pr1);
        Cell cell = row.getCell(3);
        cell.setCellValue(cell.getNumericCellValue() - pr2);       
        book.write(new FileOutputStream("D:\\1.xls"));
        book.close();                 
        InputStream ainp = new FileInputStream("D:\\1.xls");
        InputStream binp = new FileInputStream("D:\\2.xls");
        Workbook abook = new HSSFWorkbook(ainp);
        Workbook bbook = new HSSFWorkbook(binp);
        Sheet asheet = abook.getSheetAt(0);
        Row arow = asheet.getRow(pr1);
        Sheet bsheet = bbook.getSheetAt(0);
        Row brow = bsheet.getRow(pr1);        
        Cell a = arow.getCell(0);
        Cell b = arow.getCell(1);
        Cell c = arow.getCell(2);
        Cell ab = brow.createCell(0);
        Cell bb = brow.createCell(1);
        Cell cb = brow.createCell(2);
        Cell db = brow.getCell(3);
        ab.setCellValue(a.getNumericCellValue());
        bb.setCellValue(b.getStringCellValue());
        cb.setCellValue(c.getNumericCellValue());
        db.setCellValue(db.getNumericCellValue() + pr2);
        bsheet.autoSizeColumn(1);
        bsheet.autoSizeColumn(2);
        bsheet.autoSizeColumn(3);        
        bbook.write(new FileOutputStream("D:\\2.xls"));
        abook.close();
        bbook.close();
              
    }
}
class post{
    public void post() throws Exception{
        Scanner in = new Scanner(System.in);
        InputStream inp = new FileInputStream("D:\\1.xls");
        Workbook book = new HSSFWorkbook(inp);
        book.getSheetAt(0);
        Sheet sheet = book.getSheetAt(0);
        sheet.getLastRowNum();
         System.out.print("Введите ID \n ");
        int a = in.nextInt();
        String b;
        int d ;
        int c ;
        if(a<sheet.getLastRowNum())
        {
            System.out.print("Количество товара\n ");
            d = in.nextInt();
            Sheet tmp1 = book.getSheetAt(0);
            Row tmp2= tmp1.getRow(a);
          Cell tmp3 = tmp2.getCell(3);       
          tmp3.setCellValue(tmp3.getNumericCellValue() + d);
          
                                              
        }
        else {
        System.out.print("Введите Наименоваие товара \n ");
        b = in.next();
        System.out.print("Цена \n ");
        c = in.nextInt();
        System.out.print("Количество товара\n ");
         d = in.nextInt();
         Row at = sheet.createRow(a);
         Cell kol1 = at.createCell(3);
           Cell id1 = at.createCell(0);
        Cell tov1 = at.createCell(1);
        Cell price1 = at.createCell(2);       
         id1.setCellValue(a);
         tov1.setCellValue(b);
          price1.setCellValue(c);
           kol1.setCellValue(d);
        }
            sheet.autoSizeColumn(1);
            sheet.autoSizeColumn(2);
            sheet.autoSizeColumn(3);           
           book.write(new FileOutputStream("D:\\1.xls")); 
        book.close();
      
    }
}
class jur{
     public void excel() throws Exception {
        
        InputStream in = new FileInputStream("D:\\2.xls");
        HSSFWorkbook wb = new HSSFWorkbook(in);
 
        Sheet sheet = wb.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        while (it.hasNext()) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            while (cells.hasNext()) {
                Cell cell = cells.next();
                int cellType = cell.getCellType();
                switch (cellType) {
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(cell.getStringCellValue() + "||");
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print("[" + cell.getNumericCellValue() + "]");
                        break;
    
                }
            }
            System.out.println();
        }
    }
}
    