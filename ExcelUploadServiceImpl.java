package jp.co.wisdom_technology.employee.service;

import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import jp.co.wisdom_technology.employee.model.ExcelUploadModel;

/**
 * Service実装クラス
 */
@Service
public class ExcelUploadServiceImpl implements ExcelUploadService{


//	private ExcelUploadModel excelUploadModel;

	@Override
	public java.util.List<ExcelUploadModel> getBankListByExcel(InputStream in, String fileName) throws Exception {
		ArrayList<ExcelUploadModel> list = new ArrayList<ExcelUploadModel>();
        
        //创建Excel工作薄
        Workbook work= this.getWorkbook(in, fileName);
        if (null == work) {
            throw new Exception("创建Excel工作薄为空！");
        }
        Sheet sheet = null;
        Row row = null;
        Cell cell = null;
        ExcelUploadModel excelUploadModel = new ExcelUploadModel();
        for (int i = 0; i < work.getNumberOfSheets(); i++) {
            sheet = work.getSheetAt(i);
            if (sheet == null) {
                continue;
            }
            
            int attendanceDays=0;
            int daysoff_flg =0;
            double employment =0.0;
            double overtimeMorning=0.0;

            for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {
                row = sheet.getRow(j);
                
                if(row.getRowNum()==1) {
                   cell = row.getCell(3);
	               excelUploadModel.setName(cell.getStringCellValue());
	               cell = row.getCell(13);
	               excelUploadModel.setDepartment(cell.getStringCellValue());
	            }
                if(row.getRowNum()==4) {
                	cell = row.getCell(1);
 	               excelUploadModel.setYear(cell.getStringCellValue());
 	            }
                if(row.getRowNum()==5) {
                	cell = row.getCell(1);
 	               excelUploadModel.setMonth(cell.getStringCellValue());
 	            }
                
	            if(row.getRowNum()==41) {
//		                for (int y = 4; y < row.getLastCellNum(); y++) {
		                	cell = row.getCell(4);
		                	SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");  
		                	row.getCell(4).setCellType(cell.CELL_TYPE_NUMERIC);
		                	row.getCell(5).setCellType(cell.CELL_TYPE_NUMERIC);
		                	
							
		                	row.getCell(6).setCellType(cell.CELL_TYPE_NUMERIC);
		                	row.getCell(7).setCellType(cell.CELL_TYPE_STRING);
		                	
		                	row.getCell(8).setCellType(cell.CELL_TYPE_NUMERIC);
		                	row.getCell(9).setCellType(cell.CELL_TYPE_NUMERIC);
		                	row.getCell(10).setCellType(cell.CELL_TYPE_NUMERIC);
		                	row.getCell(11).setCellType(cell.CELL_TYPE_NUMERIC);
		                	row.getCell(12).setCellType(cell.CELL_TYPE_NUMERIC);
		                	row.getCell(13).setCellType(cell.CELL_TYPE_NUMERIC);
		                	row.getCell(14).setCellType(cell.CELL_TYPE_NUMERIC);
		                	row.getCell(15).setCellType(cell.CELL_TYPE_NUMERIC);
		                	row.getCell(16).setCellType(cell.CELL_TYPE_NUMERIC);
//		                	String cellValue = String.valueOf(cell.getNumericCellValue()); 
//		                	string xxx=sheet.getCellComment(4, 41).getString();
		                	excelUploadModel.setAttendanceDays(row.getCell(4).getNumericCellValue());
		                	excelUploadModel.setDaysOff(row.getCell(5).getNumericCellValue());
		                	excelUploadModel.setEmployment(sdf.format(row.getCell(6).getDateCellValue()));
		                	excelUploadModel.setOvertimeMorning(row.getCell(7).getStringCellValue());
		                	
		                	excelUploadModel.setOvertimeDay(row.getCell(8).getNumericCellValue());
		                	
		                	excelUploadModel.setOvertimeNight(row.getCell(9).getNumericCellValue());
		                	
		                	excelUploadModel.setRest(row.getCell(10).getNumericCellValue());
		                	excelUploadModel.setLate(row.getCell(11).getNumericCellValue());
		                	excelUploadModel.setAdvanceLeave(row.getCell(12).getNumericCellValue());
		                	excelUploadModel.setPaidLeave(row.getCell(13).getNumericCellValue());
		                	
		                	excelUploadModel.setSpecialLeave(row.getCell(14).getNumericCellValue());
		                	
		                	excelUploadModel.setReplaceLeave(row.getCell(15).getNumericCellValue());
		                	
		                	excelUploadModel.setUnrecognizedLeave(row.getCell(16).getNumericCellValue());
	
//		               }
                }
            }
            
            list.add(excelUploadModel);
        }
        
        work.close();
        return  list;
    }

    /**
     * 判断文件格式
     *
     * @param inStr
     * @param fileName
     * @return
     * @throws Exception 
     */
    public Workbook getWorkbook(InputStream inStr, String fileName) throws Exception {
        Workbook workbook = null;
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        if (".xls".equals(fileType)) {
            workbook = new HSSFWorkbook(inStr);
        } else if (".xlsx".equals(fileType)) {
            workbook = new XSSFWorkbook(inStr);
        } else {
            throw new Exception("请上传excel文件！");
        }
        return workbook;
    }
}
