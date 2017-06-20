package com.tjcloud.report.tenant.facade;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtils 
{
	//所传的数据为俩个list。分为 表头turpList 和 数据voList。list的数据组装格式需按照底部main函数的示例
	// 表头的英文字段需对应VO里的变量名,例：表头Turple<String, String> t = new Turple<String, String>("matchName", "赛事名称");   matchName对应matchVO里的变量 matchName
	public static final Logger logger = LoggerFactory.getLogger(ExcelUtils.class);
	public static final int SHEET_MAX_CNT = 20000; //每个sheet最多20000条数据
	
	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static void excelExport(List dataList, List<Turple<String, String>> headerList, String name, HttpServletResponse response) throws NoSuchFieldException, SecurityException, IllegalArgumentException, IllegalAccessException, UnsupportedEncodingException {
    	if(dataList == null || dataList.size() == 0 || headerList == null || headerList.size() == 0)
    	{
    		//throw Exception;
    	}
		String codedFileName = null;
		response.setContentType("application/vnd.ms-excel");
		codedFileName = java.net.URLEncoder.encode(name, "UTF-8");
		response.setHeader("content-disposition", "attachment;filename=" + codedFileName + ".xls");//告诉浏览器已下载的形式打开文件

		OutputStream fileOut = null;

		logger.info("excel导出开始--------------------");
		HSSFWorkbook wb = createExcel(dataList, headerList);
		try
		{
		//	fileOut = new FileOutputStream("E:\\matchExport.xls");
			fileOut = response.getOutputStream();
			wb.write(fileOut);
			logger.info("excel导出结束--------------------");
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		finally
		{
			if(fileOut != null)
			{
				try{
					fileOut.flush();
					fileOut.close();
				}catch(Exception e){
					e.printStackTrace();
				}
			}
		}
	}

	private static int getSheetNum(List<Object> voList) 
	{
		BigDecimal num = new BigDecimal(voList.size()).divide(new BigDecimal(SHEET_MAX_CNT+""), BigDecimal.ROUND_CEILING);
		return Integer.valueOf(num.toString());
	}

	@SuppressWarnings("rawtypes")
	public static HSSFWorkbook createExcel(List<Object> dataList, List<Turple<String, String>> headerList) throws NoSuchFieldException, SecurityException, IllegalArgumentException, IllegalAccessException
	{
		HSSFWorkbook wb = new HSSFWorkbook();
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);//设置cell的格式

		int a = 0;
		int sheetNum = getSheetNum(dataList);//获取需要多少sheet
		logger.info("总共有sheet:" + sheetNum);
		for(int s=1; s<=sheetNum; s++)
		{
			HSSFSheet sheet = wb.createSheet();
			sheet.setDefaultColumnWidth(12);
			
			List list = getHeader(headerList, sheet);
			//这里处理VOList的部分
			for(int i = a, ro = 0; i<dataList.size(); i++, ro++)//i层循环决定有多少行
			{
				HSSFRow row = sheet.createRow(ro+1);
				for(int j=0; j<list.size(); j++)//j层循环完毕是一行结束 换行
				{
					Field f= dataList.get(i).getClass().getDeclaredField(list.get(j)+"");
					f.setAccessible(true);
					HSSFCell cell = row.createCell(j);
		    		cell.setCellValue(f.get(dataList.get(i)) == null ? "" : f.get(dataList.get(i)) +"");
		    		logger.info(f.get(dataList.get(i))+"");
					//System.out.print(f.get(voList.get(i))+"----");
				}
				//System.out.println("");//数据换行
				if(i == (SHEET_MAX_CNT*s -1) )
				//if(new BigDecimal(i+"").equals(new BigDecimal(SHEET_MAX_CNT+"").multiply(new BigDecimal(s+"").subtract(new BigDecimal("1")))))
				{
					a = i+1 ;
					logger.info("新建sheet,数据坐标位" + a);
					break;
				}
			}	
		}
		return wb;
	}
	
	//获取表头的英文和汉字
	public static List<Object> getHeader(List<Turple<String, String>> headerList, HSSFSheet sheet) throws NoSuchFieldException, SecurityException, IllegalArgumentException, IllegalAccessException
    {
    	List<Object> engHeaderListList = new ArrayList<Object>();//表头英文名list
    	List<Object> chiHeadList = new ArrayList<Object>();//表头名汉字list
    	for(int i=0; i<headerList.size(); i++)
    	{
    		Field englishHead = headerList.get(i).getClass().getDeclaredField("first");
	    	Field chineseHead = headerList.get(i).getClass().getDeclaredField("second");
			englishHead.setAccessible(true);
			chineseHead.setAccessible(true);
			engHeaderListList.add(englishHead.get(headerList.get(i)));
			chiHeadList.add(chineseHead.get(headerList.get(i)));
    	}
    	HSSFRow rowHead = sheet.createRow(0);
    	for(int i=0; i<chiHeadList.size(); i++)
    	{
    		HSSFCell cell = rowHead.createCell(i);
    		cell.setCellValue(chiHeadList.get(i)+"");
    		logger.info(chiHeadList.get(i)+"");
    		//System.out.print(bList.get(i)+"--");//表头
    	}
    	//System.out.println("");//表头换行
    	return engHeaderListList;
    }
}
