package com.xyz.tt;

import javax.swing.*;

public class Test {
	
	public int getAdvice() {
        int advice = 0;
        long bfr = Math.round(26.1);
        if (bfr >= 32) {
            advice = 4;
        } else if (bfr >= 25) {
            advice = 3;
        } else if (bfr >= 21) {
            advice = 2;
        } else if (bfr >= 16) {
            advice = 1;
        }
        if (advice == 3) {
            System.out.print(advice);
        }
        return advice;
    }
	
	public static void main(String[] args) {
		//int j=new Test().getAdvice();
		System.out.println("我们是共产主义接班人");
        JOptionPane.showConfirmDialog(null, "任务完成", "Result" ,
                JOptionPane.CLOSED_OPTION);
	}


}
