package org.forlily.exceltoword;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.text.DecimalFormat;

import pl.jsolve.templ4docx.core.Docx;
import pl.jsolve.templ4docx.core.VariablePattern;
import pl.jsolve.templ4docx.variable.TextVariable;
import pl.jsolve.templ4docx.variable.Variables;

/**
 * excel to world
 *
 */
public class App {
	public static void main(String[] args) throws IOException {
		DecimalFormat df = new DecimalFormat("#,###.00");
		DecimalFormat df2 = new DecimalFormat("#####00");
		// 创建 Excel 文件的输入流对象
		FileInputStream excelFileInputStream = new FileInputStream(
				"/home/fisher/Documents/javatest/test/预付询证函控制表.xlsx");
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
		// employees.xlsx 第一行是标题行，我们从第二行开始, 对应的行索引是 1
		// sheet.getLastRowNum() : 获取当前表格中最后一行数据对应的行索引
		for (int rowIndex = 2; rowIndex <= 30; rowIndex++) {
			// XSSFRow 代表一行数据
			XSSFRow row = sheet.getRow(rowIndex);
			if (row == null) {
				continue;
			}
			XSSFCell noCell = row.getCell(0); // 序号
			XSSFCell cop_nameCell = row.getCell(1); // 客户名称
			XSSFCell ccy_nameCell = row.getCell(2); // 币种名称
			XSSFCell amtCell = row.getCell(3); // 函证金额
			
			String no=df2.format(noCell.getNumericCellValue());
			String cop_name=cop_nameCell.getStringCellValue();
			String ccy_name=ccy_nameCell.getStringCellValue();
			String amt= df.format(amtCell.getNumericCellValue());

//			StringBuilder employeeInfoBuilder = new StringBuilder();
//			employeeInfoBuilder.append("员工信息 --> ").append("序号 : ").append(noCell.getNumericCellValue())
//					.append(" , 客户名称 : ").append(cop_nameCell.getStringCellValue()).append(" , 币种 : ")
//					.append(ccy_nameCell.getStringCellValue()).append(" , 函证金额 : ")
//					.append(df.format(amtCell.getNumericCellValue()));
//			System.out.println(employeeInfoBuilder.toString());
//			
			Docx docx = new Docx("/home/fisher/Documents/javatest/test/预付询证函.docx");
			docx.setVariablePattern(new VariablePattern("#{", "}"));
			// preparing variables
			Variables variables = new Variables();
			variables.addTextVariable(new TextVariable("#{no}", no));
			variables.addTextVariable(new TextVariable("#{cop_name}", cop_name));
			variables.addTextVariable(new TextVariable("#{ccy_name}", ccy_name));
			variables.addTextVariable(new TextVariable("#{amt}", amt));
			// fill template
			docx.fillTemplate(variables);
			// save filled .docx file
			docx.save("/home/fisher/Documents/javatest/test/"+no+"_"+cop_name+".docx");
			System.out.println("/home/fisher/Documents/javatest/test/"+no+"_"+cop_name+".docx ");
			
		}
		// 操作完毕后，记得要将打开的 XSSFWorkbook 关闭
		workbook.close(); // （注意：所有操作完毕后，统一关闭，如果后面还有关于这个Excel文件的操作，这里先不关闭，但所有操作完成后，一定记得关闭对象！）
		System.out.println("Done!");
	}



}
