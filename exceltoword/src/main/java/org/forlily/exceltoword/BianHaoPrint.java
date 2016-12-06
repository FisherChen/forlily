package org.forlily.exceltoword;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import pl.jsolve.templ4docx.core.Docx;
import pl.jsolve.templ4docx.core.VariablePattern;
import pl.jsolve.templ4docx.variable.TextVariable;
import pl.jsolve.templ4docx.variable.Variables;

public class BianHaoPrint {
	
	public static void main(String[] args) throws IOException {

		//第一步 约定Excel存放的位置，并指定说明，程序需要从excel的第多少行读取到第多少行。
		String Excel_File_Path="/home/fisher/Documents/javatest/test/BianHaoPrint/model/多少个Word.xlsx";
		// 指定Word的模板存放位置 ,注意Word目前仅只能是“docx”的格式
		String WordModel_File_Path="/home/fisher/Documents/javatest/test/BianHaoPrint/model/路灯调查问卷模板.docx";
		// 指定生成的文件存放的文件夹位置
		String Save_Word_File_Path="/home/fisher/Documents/javatest/test/BianHaoPrint/";
		// 文件的名称：
		String doc_name_prix="路灯调查问卷";
		// 开始行数
		int Start_Row_Index_No=2;
		// 结束行行数
		int END_Row_Index_No=101;
		
		//=======================================基本不会变===========================================================================
		//第二步 约定模板中各类数据的格式
		//格式化金额，将excel中的金额格式化成用逗号分割，且有2位的小数点。需要注意excel本身金额必须是2位小数
		DecimalFormat df = new DecimalFormat("#,##0.00");
		//格式化索引，将Excel 中“序号”的第一列，格式化成至少2位数字的格式，如 1 就是01。需要注意excel本身“序号”必须是整数
		DecimalFormat df2 = new DecimalFormat("#####00");
		// 格式化日期，将Excel中 “日期”格式化成中文的“年-月-日”的格式。需要注意excel本身“日期”相关字段必须是日期类型的格式
		SimpleDateFormat formate= new SimpleDateFormat("yyyy年MM月dd日");
		
		//=======================================不需要理解开始========================================================================		
		// 创建 Excel 文件的输入流对象
		FileInputStream excelFileInputStream = new FileInputStream(Excel_File_Path);
		// XSSFWorkbook 就代表一个 Excel 文件
		// 创建其对象，就打开这个 Excel 文件
		XSSFWorkbook workbook = new XSSFWorkbook(excelFileInputStream);
		// 输入流使用后，及时关闭！这是文件流操作中极好的一个习惯！
		excelFileInputStream.close();
		// XSSFSheet 代表 Excel 文件中的一张表格
		// 我们通过 getSheetAt(0) 指定表格索引来获取对应表格
		// 注意表格索引从 0 开始！
		XSSFSheet sheet = workbook.getSheetAt(0);

		// 开始循环表格数据,表格的行索引从 0 开始
		// 预付询证函控制表.xlsx 第一行是大标题，第二行是表头, 所以从第三行开始（对应的行索引是 2）(因索引是0开始的，所以第三是2)
		// 因为
		// sheet.getLastRowNum() : 获取当前表格中最后一行数据对应的行索引
		
		for (int rowIndex = (Start_Row_Index_No-1); rowIndex <= (END_Row_Index_No-1); rowIndex++) {
			// XSSFRow 代表一行数据
			XSSFRow row = sheet.getRow(rowIndex);
			if (row == null) {
				continue;
			}
		//=======================================不需要理解结束===========================================================================		
			
			
			//===============================解析一行数据，注意顺序是从0开始=================================================================
			XSSFCell noCell = row.getCell(0); // 序号
			
			//================================不需要理解=================================================================================
			String no=df2.format(noCell.getNumericCellValue());

			
			Docx docx = new Docx(WordModel_File_Path);
			docx.setVariablePattern(new VariablePattern("#{", "}"));
			// preparing variables
			Variables variables = new Variables();
			//================================不需要理解====================================================================================
			
			
			//=============================================替换Word中的变量，注意名称=========================================================
			variables.addTextVariable(new TextVariable("#{num}", no));
			
			//=============================================替换Word中的变量，注意名称=========================================================
			
			
			//================================不需要理解====================================================================================
			// fill template
			docx.fillTemplate(variables);
			// save filled .docx file
			docx.save(Save_Word_File_Path+no+"_"+doc_name_prix+".docx");
			System.out.println(Save_Word_File_Path+no+"_"+doc_name_prix+".docx");
		}
		// 操作完毕后，记得要将打开的 XSSFWorkbook 关闭
		workbook.close(); // （注意：所有操作完毕后，统一关闭，如果后面还有关于这个Excel文件的操作，这里先不关闭，但所有操作完成后，一定记得关闭对象！）
		System.out.println("Done!");
			//================================不需要理解====================================================================================
	}




}
	
