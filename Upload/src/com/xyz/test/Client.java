package com.xyz.test;

import java.util.HashMap;
import java.util.Map;

import javax.swing.JOptionPane;

import com.xyz.network.HttpUtil;
import com.xyz.network.Utility;

public class Client {

	static void getConnection() {
		HttpUtil.getConnection(new Utility.GetConnectionCallBack() {

			@Override
			public void connectionError() {
				System.out.println("CONNECTION_ERROR");
			}
		});
	}

	static void getCheckCode() {
		HttpUtil.getCheckcode(new Utility.GetCheckcodeCallBack() {
			@Override
			public void doSuccess() {
				System.out.println("成功获取验证码图片");
				String inputValue = JOptionPane.showInputDialog("Please input a value");
				login(inputValue);
			}

			@Override
			public void doError() {
				System.out.println("LOAD_ERROR");
			}
		});
	}

	static void login(String checkCode) {
		final Map<String, String> data = new HashMap<>();
		data.put("txtUserName", "00201091");
		data.put("TextBox2", "dragon7687");
		data.put("txtSecretCode", checkCode);
		System.out.println("登录中...");
		HttpUtil.login(data, new Utility.LoginCallBack() {
			@Override
			public void loginSuccess() {
				
				System.out.println("loginSuccess");
				HttpUtil.saveScoresOfCourse();
			}

			@Override
			public void loginError() {
				System.out.println("loginError");

			}
		});
	}

	public static void main(String[] args) {
		getConnection();
		getCheckCode();

	}
}
