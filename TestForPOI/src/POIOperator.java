import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * Created by finnflyer on 2015/12/14.
 */

public class POIOperator {
    public static String outputFile = "D://Test.xls";
    public static void main(String args[]) {
        try {
            // 创建新的Excel 工作簿
            HSSFWorkbook workbook = new HSSFWorkbook();
            // 在Excel工作簿中建一工作表，其名为缺省值
            // 如要新建一名为"效益指标"的工作表，其语句为：
            HSSFSheet sheet = workbook.createSheet("Sheet1");
            // HSSFSheet sheet = workbook.createSheet();
            // 在索引0的位置创建行（最顶端的行）
            HSSFRow row = sheet.createRow(0);
            // 在索引0的位置创建单元格（左上端）
            HSSFCell cell = row.createCell(0);
            // 定义单元格为字符串类型
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            // 在单元格中输入一些内容
            cell.setCellValue(" number 1");
            HSSFCell cell01 =row.createCell(1); // row.createCell((short)1);

            // 定义单元格为字符串类型
            cell01.setCellType(HSSFCell.CELL_TYPE_STRING);
            // 在单元格中输入一些内容
            cell01.setCellValue(" number2");


            // 新建一输出文件流
            FileOutputStream fOut = new FileOutputStream(outputFile);
            // 把相应的Excel 工作簿存盘
            workbook.write(fOut);
            fOut.flush();
            // 操作结束，关闭文件
            fOut.close();
            System.out.println("文件生成...");
        } catch (Exception e) {
            System.out.println(" xlCreate() : " + e);
        }
    }

    private void ReadXls() throws IOException {
        InputStream is = new FileInputStream( "D:\\excel\\xls_test2.xls");
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook( is);

        // 循环工作表Sheet
        for(int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++){
            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt( numSheet);
            if(hssfSheet == null){
                continue;
            }

            // 循环行Row
            for(int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++){
                HSSFRow hssfRow = hssfSheet.getRow( rowNum);
                if(hssfRow == null){
                    continue;
                }

                // 循环列Cell
                for(int cellNum = 0; cellNum <= hssfRow.getLastCellNum(); cellNum++){
                    HSSFCell hssfCell = hssfRow.getCell( cellNum);
                    if(hssfCell == null){
                        continue;
                    }

                    System.out.print("    " + getValue( hssfCell));
                }
                System.out.println();
            }
        }
    }
    @SuppressWarnings("static-access")
    private String getValue(HSSFCell hssfCell){
        if(hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN){
            return String.valueOf( hssfCell.getBooleanCellValue());
        }else if(hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC){
            return String.valueOf( hssfCell.getNumericCellValue());
        }else{
            return String.valueOf( hssfCell.getStringCellValue());
        }
    }



}
