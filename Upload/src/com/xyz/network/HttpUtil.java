package com.xyz.network;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import javax.swing.JOptionPane;

/**
 * Created by zheng on 16/12/12.
 */
public class HttpUtil {

	private final static String url_base = "http://jw1.wucc.cn";
	//private final static String url_base = "http://jw1.wucc.cn:2222";
	// 主页面网址
	private final static String url_index = url_base + "/default2.aspx";

	// 验证码网址
	private final static String url_checkcode = url_base + "/CheckCode.aspx";

	// 存放cookie
	private static Map<String, String> cookies = null;

	// 存放viewstate
	private static String viewState = null;

	// 学生学号
	private static String xsxh;

	// 学生姓名
	private static String xsxm;

	private static HttpUtil httpUtil;
	private static Thread connection;

	private HttpUtil() {
	}

	// 获得HttpUtil实例
	private static HttpUtil getInstance() {
		if (httpUtil == null) {
			synchronized (HttpUtil.class) {
				if (httpUtil == null) {
					httpUtil = new HttpUtil();
				}
			}
		}
		return httpUtil;
	}

	// 获得连接，并保存cookie和viewstate在内存中
	public static void getConnection(final Utility.GetConnectionCallBack callback) {
		connection = new Thread() {
			@Override
			public void run() {
				super.run();
				getInstance()._getConnection(callback);
			}
		};
		connection.start();
	}

	// 登陆
	public static void login(final Map<String, String> data, final Utility.LoginCallBack callback) {
		new Thread() {
			@Override
			public void run() {
				super.run();
				getInstance()._login(data, callback);
			}
		}.start();
	}

	// 连接后，下载验证码
	public static void getCheckCode(final Utility.GetCheckcodeCallBack callback) {
		new Thread() {
			@Override
			public void run() {
				try {
					connection.join();
				} catch (InterruptedException e) {
					e.printStackTrace();
				}
				System.out.println("get Check code执行");
				getInstance()._getCheckcode(callback);

			}
		}.start();
	}

	// 获得连接，并保存cookie和viewState在内存中
	private void _getConnection(Utility.GetConnectionCallBack callback) {
		try {
			Connection conn = Jsoup.connect(url_index).timeout(20000).method(Connection.Method.GET);
			Connection.Response response = conn.execute();
			if (response.statusCode() == 200) {
				cookies = response.cookies();
				System.out.println("cookies = " + cookies.toString());
			} else {
				callback.connectionError();
			}
			Document document = Jsoup.parse(response.body());
			viewState = document.select("input").first().attr("value");

			System.out.println("viewState = " + viewState);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	// 下载验证码并调用回调函数。若成功，则显示验证码，否则提示错误
	private void _getCheckcode(final Utility.GetCheckcodeCallBack callback) {
		try {
			Connection.Response resultImageResponse = Jsoup.connect(url_checkcode).cookies(cookies)
					.ignoreContentType(true).execute();
			File file = new File("file", "checkCode.png");

			FileOutputStream fos = new FileOutputStream(file);
			fos.write(resultImageResponse.bodyAsBytes());
			fos.flush();
			fos.close();
			callback.doSuccess();
		} catch (Exception e) {
			callback.doError();
			e.printStackTrace();
		}
	}

	// post发送数据登录,并储存教师的姓名、教工号
	private void _login(final Map<String, String> data, final Utility.LoginCallBack callback) {
		data.put("__VIEWSTATE", viewState);
		data.put("RadioButtonList1", "教师");
		data.put("Button1", "");
		Connection conn = Jsoup.connect(url_index).cookies(cookies).header("Origin", "http://jw1.wucc.cn:2222")
				.data(data).method(Connection.Method.POST).timeout(20000);
		try {
			Connection.Response response = conn.execute();
			if (Utility.isLoginSuccess(response)) {
				xsxh = data.get("txtUserName");
				xsxm = Utility.getStudentName(response);
				System.out.println("教师工号: " + xsxh + ", 教师姓名: " + xsxm);
				Document document = Jsoup.parse(response.body());
				viewState = document.select("input").get(2).attr("value");

				System.out.println("tea viewState = " + viewState);

				callback.loginSuccess();
			} else {
				callback.loginError();
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void saveScoresOfCourse() {
		new Thread() {
			@Override
			public void run() {

				getInstance().clickCourseLink();
				//getInstance()._getAllCoursesMSG();
				try {
					Thread.sleep(3000);
				} catch (Exception e) {
					// TODO: handle exception
				}
				getInstance()._enterTheCourse();
				try {
					Thread.sleep(3000);
				} catch (Exception e) {
					// TODO: handle exception
				}
				getInstance()._saveScores();

			}
		}.start();
	}
	
	public void getAllCoursesMSG(){
		_getAllCoursesMSG();
	}

	private void _getAllCoursesMSG() {

		System.out.println("开始获取课程：");
		Connection conn = Jsoup.connect(url_base + "/js_cjcd.aspx?xh=00201091&gnmkdm=N1221").cookies(cookies)
				.header("Origin", "http://jw1.wucc.cn:2222").method(Connection.Method.GET).timeout(20000);
		try {
			Connection.Response response = conn.execute();
			//System.out.println("body:" + response.body());
			Document document = Jsoup.parse(response.body());
			viewState = document.select("input").get(0).attr("value");
			String link1 = document.select("a").get(0).attr("href");
			String link2 = document.select("a").get(1).attr("href");
			String link3 = document.select("a").get(2).attr("href");
			String link4 = document.select("a").get(3).attr("href");
			String link5 = document.select("a").get(4).attr("href");
			String title1 = document.select("a").get(0).attr("title");
			String title2 = document.select("a").get(1).attr("title");
			String title3 = document.select("a").get(2).attr("title");
			String title4 = document.select("a").get(3).attr("title");
			String title5 = document.select("a").get(4).attr("title");
			System.out.println("title1:" + title1);
			System.out.println("title2:" + title2);
			System.out.println("title3:" + title3);
			System.out.println("title4:" + title4);
			System.out.println("title5:" + title5);
			System.out.println("link1:" + link1);
			System.out.println("link2:" + link2);
			System.out.println("link3:" + link3);
			System.out.println("link4:" + link4);
			System.out.println("link5:" + link5);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void clickCourseLink() {

		System.out.println("点击链接：");
		Connection conn = Jsoup
				.connect(url_base + "/js_cjmm.aspx?xkkh="+serial+"&zgh=00201091&gnmkdm=N12211")
				.cookies(cookies).header("Origin", "http://jw1.wucc.cn:2222").method(Connection.Method.GET)
				.timeout(20000);
		try {
			Connection.Response response = conn.execute();
			//System.out.println("body:" + response.body());
			Document document = Jsoup.parse(response.body());
			viewState = document.select("input").get(0).attr("value");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private String serial;
	private String myViewState;
	private String courseId;

	private void _enterTheCourse() {
		System.out.println("开始进入课程：");
		courseId = JOptionPane.showInputDialog("Please input a Course ID");
		Map<String, String> datas = new HashMap<String, String>();

		datas.put("__VIEWSTATE",
				viewState);
		datas.put("hidLanguage:", "");

		datas.put("txtzgh", "00201091");

		datas.put("Button1", "确定");
		Sheet sheet = null;
		Workbook wb = null;
		try {

			FileInputStream fis = new FileInputStream("UploadFile/course.xlsx");
			wb = new XSSFWorkbook(fis);
			sheet = wb.getSheetAt(0);
		} catch (Exception exception) {

		}
		XSSFRow hssfrow = (XSSFRow) sheet.getRow(Integer.parseInt(courseId));
		XSSFCell pwCell = hssfrow.getCell(3);
		XSSFCell serialCell = hssfrow.getCell(4);
		String pw = String.valueOf((int) pwCell.getNumericCellValue());
		System.out.println("pw:" + pw);
		serial = serialCell.getStringCellValue();
		System.out.println("serial" + serial);
		datas.put("TextBox1", pw);
		datas.put("txtxkkh", serial);
		String url = url_base + "/js_cjmm.aspx?xkkh=" + serial + "&zgh=00201091&gnmkdm=N12211";
		Connection conn = Jsoup.connect(url).cookies(cookies).header("Origin", "http://jw1.wucc.cn:2222").data(datas)
				.method(Connection.Method.POST).timeout(20000);
		try {
			Connection.Response response = conn.execute();
			String dest = new String(response.bodyAsBytes(), "gb2312");
			Utility.writeTxtFile(dest, "UploadFile/course" + courseId + ".html");
			Document document = Jsoup.parse(dest);
			myViewState = document.select("input").get(2).attr("value");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void _saveScores() {
		System.out.println("开始提交课程成绩：");
		Map<String, String> datas = new HashMap<String, String>();

		datas.put("__VIEWSTATE", myViewState);
		datas.put("hidLanguage:", "");
		datas.put("txtxkkh", serial);
		datas.put("sfts", "true");
		datas.put("txtzgh", "00201091");
		datas.put("TextBox1", "");
		datas.put("ch", "1");
		datas.put("jfz", "百分制");
		datas.put("psb", "40");// 平时百分比
		datas.put("qzb", "0");// 期中百分比
		datas.put("qmb", "60");// 期末百分比
		datas.put("syb", "0");// 实验百分比
		datas.put("txtChanged", "0");
		datas.put("Dropdownlist1", "百分制");

		FileInputStream fis;
		Workbook wb = null;
		try {
			fis = new FileInputStream("UploadFile/course" + courseId + ".xlsx");
			wb = new XSSFWorkbook(fis);
		} catch (IOException e1) {
			e1.printStackTrace();
		}

		Sheet sheet = wb.getSheetAt(0);
		int stuCount = sheet.getLastRowNum() - 2;
		for (int i = 0; i <= stuCount; i++) {
			int row = i + 2;
			String pre = "DataGrid1:_ctl";
			String ps2 = ":ps";
			String qm2 = ":qm";
			XSSFRow hssfrow = (XSSFRow) sheet.getRow(row);
			XSSFCell psCell = hssfrow.getCell(2);
			XSSFCell qmCell = hssfrow.getCell(3);
			String ps = String.valueOf((int) psCell.getNumericCellValue());
			String qm = String.valueOf((int) qmCell.getNumericCellValue());
			datas.put(pre + row + ps2, ps);
			datas.put(pre + row + qm2, qm);
			datas.put(pre + row + ":dbz", "");
			datas.put(pre + row + ":qz", "");
			datas.put(pre + row + ":sy", "");
			datas.put(pre + row + ":zp", "");
		}
		datas.put("Button1", "保存");

		Iterator<Map.Entry<String, String>> entries = datas.entrySet().iterator();

		while (entries.hasNext()) {

			Map.Entry<String, String> entry = entries.next();

			System.out.println(entry.getKey() + "=" + entry.getValue());

		}
		String url = url_base + "/xf_js_cjlr.aspx?xkkh=" + serial + "&zgh=00201091&gnmkdm=N12211";

		Connection conn = Jsoup.connect(url).cookies(cookies).header("Origin", "http://jw1.wucc.cn:2222").data(datas)
				.method(Connection.Method.POST).timeout(20000);
		try {
			Connection.Response response = conn.execute();

			String dest = new String(response.bodyAsBytes(), "gb2312");
			Utility.writeTxtFile(dest, "UploadFile/scores" + courseId + ".html");

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
