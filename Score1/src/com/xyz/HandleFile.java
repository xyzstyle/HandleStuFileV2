package com.xyz;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class HandleFile {

    private String filePath;
    private int nameCol = 1;

    public HandleFile(String filePath) {
        this.filePath = filePath;
    }

    public static String[] readFiles(String filePath)
            throws FileNotFoundException, IOException {
        File file = new File(filePath);
        if (file.isDirectory()) {
            //System.out.println("文件夹");
            String[] filelist = file.list();
            // for (int i = 0; i < filelist.length; i++) {
            // File readfile = new File(filePath + "\\" + filelist[i]);
            // if (!readfile.isDirectory()) {
            // System.out.println("path=" + readfile.getPath());
            // System.out.println("absolutepath="
            // + readfile.getAbsolutePath());
            // System.out.println("name=" + readfile.getName());
            //
            // } else if (readfile.isDirectory()) {
            // // readfile(filePath + "\\" + filelist[i]);
            // }
            // }
            return filelist;

        }
        return null;
    }

    public ArrayList<String> getNames() {
        String filename = filePath + "/labs.xls";
        ArrayList<String> nameList = null;
        try {
            FileInputStream file = new FileInputStream(filename);
            POIFSFileSystem ts = new POIFSFileSystem(file);
            HSSFWorkbook wb = new HSSFWorkbook(ts);
            HSSFSheet sh = wb.getSheetAt(0);
            System.out.println(wb.getSheetName(0));
            HSSFRow ro = null;

            for (int i = 0; sh.getRow(i) != null; i++) {
                ro = sh.getRow(i);

                for (int j = 0; ro.getCell(j) != null; j++) {
                    if (ro.getCell(j).getStringCellValue().equals("姓名")) {
                        nameCol = j;
                        break;
                    }
                }
            }
            nameList = new ArrayList<String>();
            for (int i = 1; sh.getRow(i) != null; i++) {
                ro = sh.getRow(i);
                HSSFCell cell = ro.getCell(1);
                if (cell != null) {
                    nameList.add(cell.getStringCellValue());
                }

            }
            System.out.println(nameList);
            wb.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return nameList;
    }

    public ArrayList<String> getLabs() {
        ArrayList<String> labList = null;
        String filename = filePath + "/labs.xls";
        try {
            FileInputStream file = new FileInputStream(filename);
            POIFSFileSystem ts = new POIFSFileSystem(file);
            HSSFWorkbook wb = new HSSFWorkbook(ts);
            HSSFSheet sh = wb.getSheetAt(0);
            HSSFRow ro = sh.getRow(0);
            labList = new ArrayList<String>();
            for (int i = 2; ro.getCell(i) != null; i++) {
                HSSFCell cell = ro.getCell(i);
                if (cell != null) {
                    labList.add(cell.getStringCellValue());
                }
            }
            System.out.println(labList);
            wb.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return labList;
    }

    public static boolean isHave(String[] strs, String s) {
        /*
		 * 此方法有两个参数，第一个是要查找的字符串数组，第二个是要查找的字符或字符串
		 */
        for (int i = 0; i < strs.length; i++) {
            if (strs[i].indexOf(s) != -1) {
                return true;
            }
        }

        return false;// 没找到返回false
    }

    public ArrayList<HashMap<String, Integer>> getBlanks(
            ArrayList<String> nameList, ArrayList<String> labList) {

        if (nameList == null || labList == null) {
            return null;
        }
        ArrayList<HashMap<String, Integer>> blankList = new ArrayList<HashMap<String, Integer>>();
        for (int i = 0; i < labList.size(); i++) {
            String currentLab = labList.get(i);
            String[] files = null;
            try {
                files = readFiles(filePath + "/" + currentLab);

            } catch (FileNotFoundException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
            for (int ii = 0; ii < files.length; ii++) {
                System.out.println(files[ii]);
            }
            for (int j = 0; j < nameList.size(); j++) {
                if (!isHave(files, nameList.get(j))) {
                    HashMap<String, Integer> theBlack = new HashMap<String, Integer>();
                    theBlack.put("row", j);
                    theBlack.put("col", i);
                    blankList.add(theBlack);

                }

            }

        }
        System.out.println(blankList);
        return blankList;

    }

    public void handle(ArrayList<HashMap<String, Integer>> blankList) {

        try {
            FileInputStream file = new FileInputStream("file1/labs.xls");
            POIFSFileSystem ts = new POIFSFileSystem(file);
            HSSFWorkbook wb = new HSSFWorkbook(ts);
            HSSFSheet sh = wb.getSheetAt(0);
            HSSFRow ro = null;
            for (HashMap<String, Integer> theBlank : blankList) {

                ro = sh.getRow(theBlank.get("row") + 1);
                HSSFCell cell = ro.getCell(theBlank.get("col") + nameCol + 1);
                cell.setCellValue("");
            }

            FileOutputStream out = null;

            try {

                out = new FileOutputStream("file1/labs111.xls");

                wb.write(out);

            } catch (IOException e) {

                e.printStackTrace();

            } finally {

                try {

                    out.close();

                } catch (IOException e) {

                    e.printStackTrace();

                }

            }
            wb.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public static void main(String[] args) {

        ArrayList<String> nameList, labList;
        ArrayList<HashMap<String, Integer>> blankList;
        HandleFile handleFile = new HandleFile("file1");
        nameList = handleFile.getNames();
        labList = handleFile.getLabs();
        blankList = handleFile.getBlanks(nameList, labList);
        handleFile.handle(blankList);
        JOptionPane.showConfirmDialog(null, "Mission accomplished,Please check results", "Result" ,
                JOptionPane.CLOSED_OPTION);


    }

}
