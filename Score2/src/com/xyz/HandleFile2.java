package com.xyz;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.*;

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 处理学生平时成绩的静态类
 * 
 * @author xyz
 *
 */
public  class HandleFile2 {

	/**
	 * 
	 * @param filePath
	 *            目录
	 * @return String[] 目录下所有文件的字符数组
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static String[] readFileNamesOfPath(String filePath) throws FileNotFoundException, IOException {
		File file = new File(filePath);
		if (file.isDirectory()) {
			String[] filelist = file.list();
			for (String s : filelist) {
				System.out.println(s);
			}
			return filelist;
		}
		return null;
	}

	public static void gatherStuScore() {
		String[] files = null;
		try {
			files = readFileNamesOfPath("file2/stu");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		if (files == null) {
			JOptionPane.showConfirmDialog(null, "The lack of students score file under the dir", "Error", JOptionPane.CLOSED_OPTION);
			return;
		}
		Sheet sheet = null;
		Workbook wb = null;
		try {

			FileInputStream fis = new FileInputStream("file2/all.xlsx");
			wb = new XSSFWorkbook(fis);
			sheet = wb.getSheetAt(0);
		} catch (Exception exception) {

		}
		if (sheet == null) {
			JOptionPane.showConfirmDialog(null, "缺乏汇总文件", "错误", JOptionPane.CLOSED_OPTION);
			return;
		}

		for (String file : files) {
			copyScores(sheet, file);
		}

		FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream("file2/all.xlsx");
			wb.write(fileOut);
			fileOut.close();
		} catch (IOException e) {

			e.printStackTrace();
		}

	}

	private static void copyScores(Sheet gatherSheet, String file) {
		String fileName = file.split("\\.")[0];
		System.out.println("fileName:" + fileName);
		double[] scores = null;
		try {
			FileInputStream fis = new FileInputStream("file2/stu/" + file);
			Workbook wb = new XSSFWorkbook(fis);
			Sheet sheet = wb.getSheetAt(0);
			int rowCounts = sheet.getLastRowNum();
			FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
			scores = new double[32];
			System.out.println("rowcounts:" + rowCounts);
			for (int i = 2; i < 34; i++) {
				XSSFRow hssfrow = (XSSFRow) sheet.getRow(i);
				XSSFCell scoreCell = hssfrow.getCell(5);//cellnum为数据源的分数所在的列号，即学生所打的分数
				evaluator.evaluateFormulaCell(scoreCell);
				scores[i - 2] = scoreCell.getNumericCellValue();
				System.out.println(scoreCell.getNumericCellValue());
				XSSFRow hssfrow1 = (XSSFRow) gatherSheet.getRow(i);
				XSSFCell scoreCell1 = hssfrow1.getCell(Integer.parseInt(fileName) + 3);//目的汇总文件中该分数所应保存的列号
				scoreCell1.setCellValue(scores[i - 2]);
			}
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	/**
	 * 分组之后的姓名顺序和之前的姓名顺序是不一样的
	 * 将分组之后的分数从soft1401.xlsx正确拷贝到labs.xlsx
	 * 因为用姓名作为比较项，因此对于一个班级之中有同姓名的同学可能产生错误的分数拷贝
	 */
	public static void getScore() {

		try {
			FileInputStream isDestination = new FileInputStream("file2/Soft1402.xlsx");
			Workbook wbDestination = new XSSFWorkbook(isDestination);
			Sheet sheetDestination = wbDestination.getSheetAt(0);
			int rowCountsDestination = sheetDestination.getLastRowNum();
			String[] stuNames = new String[rowCountsDestination-1];

			FileInputStream isSource = new FileInputStream("file2/all.xlsx");
			Workbook wbSource = new XSSFWorkbook(isSource);
			Sheet sheetSource = wbSource.getSheetAt(0);
			FormulaEvaluator evaluator = wbSource.getCreationHelper().createFormulaEvaluator();

			for (int i = 2; i <= rowCountsDestination; i++) {
				XSSFRow hssfrowDestination = (XSSFRow) sheetDestination.getRow(i);
				XSSFCell nameCellDestination = hssfrowDestination.getCell(1);
				stuNames[i - 2] = nameCellDestination.getStringCellValue().trim();
				XSSFCell scoreCellDestination = hssfrowDestination.getCell(2);
				System.out.println();
				System.out.println("i:"+i);
				for (int j = 2; j <= rowCountsDestination; j++) {
				    System.out.print("J:"+j);
					XSSFRow hssfrowSource = (XSSFRow) sheetSource.getRow(j);
					XSSFCell nameCellSource = hssfrowSource.getCell(3);
					XSSFCell scoreCellSource = hssfrowSource.getCell(21);
					if (stuNames[i - 2].equals(nameCellSource.getStringCellValue().trim())) {
						evaluator.evaluateFormulaCell(scoreCellSource);
						scoreCellDestination.setCellValue((int) scoreCellSource.getNumericCellValue());
						break;
					}
				}
			}
			FileOutputStream fileOut;
			fileOut = new FileOutputStream("file2/score.xlsx");
			wbDestination.write(fileOut);
			fileOut.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public static void main(String[] args) {

	    JFrame frame= new JFrame("Score");
	    frame.setLocation(500,300);
        frame.setSize(new Dimension(400,200));
        frame.setLayout(new FlowLayout());

        JButton gatherScoreBtn = new JButton("Gather Score");
        gatherScoreBtn.setPreferredSize(new Dimension(144,35));
        Font f=new Font(null,Font.BOLD,16);
        gatherScoreBtn.setFont(f);
        gatherScoreBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                gatherStuScore();
            }
        });
        frame.add(gatherScoreBtn);
        JButton getScoreBtn = new JButton("Get Score");
        getScoreBtn.setPreferredSize(new Dimension(144,35));
        getScoreBtn.setFont(f);
        getScoreBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                getScore();
            }
        });
        frame.add(getScoreBtn);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setVisible(true);

	}

}
