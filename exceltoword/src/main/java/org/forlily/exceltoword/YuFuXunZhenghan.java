package org.forlily.exceltoword;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;

import pl.jsolve.templ4docx.core.Docx;
import pl.jsolve.templ4docx.core.VariablePattern;
import pl.jsolve.templ4docx.variable.TextVariable;
import pl.jsolve.templ4docx.variable.Variables;

/**
 * excel to world
 *
 */
public class YuFuXunZhenghan {
	public static void main(String[] args) throws IOException {

		//第一步 约定Excel存放的位置，并指定说明，程序需要从excel的第多少行读取到第多少行。
		String Excel_File_Path="/home/fisher/Documents/javatest/test/预付询证函控制表.xlsx";
		// 指定Word的模板存放位置 ,注意Word目前仅只能是“docx”的格式
		String WordModel_File_Path="/home/fisher/Documents/javatest/test/预付询证函.docx";
		// 指定生成的文件存放的文件夹位置
		String Save_Word_File_Path="/home/fisher/Documents/javatest/test/";
		// 开始行数
		int Start_Row_Index_No=3;
		// 结束行行数
		int END_Row_Index_No=31;
		
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
			XSSFCell cop_nameCell = row.getCell(1); // 客户名称
			XSSFCell ccy_nameCell = row.getCell(2); // 币种名称（欠贵公司）
			XSSFCell amtCell = row.getCell(3); // 函证金额（欠贵公司）
			XSSFCell ccy_nameCell2 = row.getCell(4); // 币种名称（贵公司欠）
			XSSFCell amtCell2 = row.getCell(5); // 函证金额（贵公司欠）
			XSSFCell endDateCell = row.getCell(6); // 截至日期
			XSSFCell beforeRemarksCell = row.getCell(7); // 去函备注
			
			//================================不需要理解=================================================================================
			String no=df2.format(noCell.getNumericCellValue());
			String cop_name=cop_nameCell.getStringCellValue();
			String ccy_name=ccy_nameCell.getStringCellValue();
			String amt= df.format(amtCell.getNumericCellValue());
			String ccy_name2=ccy_nameCell2.getStringCellValue();
			String amt2= df.format(amtCell2.getNumericCellValue());
			String endDate= formate.format(endDateCell.getDateCellValue());
			String beforeRemarks=beforeRemarksCell.getStringCellValue();

			/**
			StringBuilder employeeInfoBuilder = new StringBuilder();
			employeeInfoBuilder.append("员工信息 --> ")
					.append("序号 : ").append(df2.format(noCell.getNumericCellValue()))
					.append(" , 客户名称 : ").append(cop_nameCell.getStringCellValue())
					.append(" , 币种（欠贵公司） : ").append(ccy_nameCell.getStringCellValue())
					.append(" , 函证金额 （欠贵公司）: ")	.append(df.format(amtCell.getNumericCellValue()))
					.append(" , 币种（贵公司欠） : ").append(ccy_nameCell2.getStringCellValue())
					.append(" , 函证金额 （贵公司欠）: ")	.append(df.format(amtCell2.getNumericCellValue()))
					.append(" , 截至日期: ").append(formate.format(endDateCell.getDateCellValue()))
					.append(" , 去函备注 ").append(beforeRemarksCell.getStringCellValue())
					;
			System.out.println(employeeInfoBuilder.toString());
			**/
			
			Docx docx = new Docx(WordModel_File_Path);
			docx.setVariablePattern(new VariablePattern("#{", "}"));
			// preparing variables
			Variables variables = new Variables();
			//================================不需要理解====================================================================================
			
			
			//=============================================替换Word中的变量，注意名称=========================================================
			variables.addTextVariable(new TextVariable("#{no}", no));
			variables.addTextVariable(new TextVariable("#{cop_name}", cop_name));
			variables.addTextVariable(new TextVariable("#{ccy_name}", ccy_name));
			variables.addTextVariable(new TextVariable("#{amt}", amt));
			variables.addTextVariable(new TextVariable("#{ccy_name2}", ccy_name2));
			variables.addTextVariable(new TextVariable("#{amt2}", amt2));
			variables.addTextVariable(new TextVariable("#{endDate}", endDate));
			variables.addTextVariable(new TextVariable("#{beforeRemarks}", beforeRemarks));
			
			//=============================================替换Word中的变量，注意名称=========================================================
			
			
			//================================不需要理解====================================================================================
			// fill template
			docx.fillTemplate(variables);
			// save filled .docx file
			docx.save(Save_Word_File_Path+no+"_"+cop_name+".docx");
			System.out.println(Save_Word_File_Path+no+"_"+cop_name+".docx");
		}
		// 操作完毕后，记得要将打开的 XSSFWorkbook 关闭
		workbook.close(); // （注意：所有操作完毕后，统一关闭，如果后面还有关于这个Excel文件的操作，这里先不关闭，但所有操作完成后，一定记得关闭对象！）
		System.out.println("Done!");
			//================================不需要理解====================================================================================
	}



}
