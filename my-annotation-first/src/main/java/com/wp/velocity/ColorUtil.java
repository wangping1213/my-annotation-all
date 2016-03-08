package com.wp.velocity;

import org.apache.commons.lang.StringUtils;

import java.awt.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 颜色转换工具
 * 类名称：ColorUtil
 * 创建人：wangping
 * 修改人：wangping
 * 修改时间： 2015年4月7日 下午3:11:04
 */
public class ColorUtil {

	/**
	 * 将color对象转换为css模式的颜色字符串（非简写模式 #0066cc）
	 * @param color 颜色对象
	 * @return 返回css模式的颜色字符串
	 * @exception
	 */
	public static String Color2String(Color color) {
		String R = Integer.toHexString(color.getRed());
		R = R.length() < 2 ? ('0' + R) : R;
		String B = Integer.toHexString(color.getBlue());
		B = B.length() < 2 ? ('0' + B) : B;
		String G = Integer.toHexString(color.getGreen());
		G = G.length() < 2 ? ('0' + G) : G;
		return '#' + R + B + G;
	}

	/**
	 * 将css模式的颜色字符串转换为Color对象
	 * @param str css模式的颜色字符串
	 * @return 返回color对象
	 * @exception
	 */
	public static Color String2Color(String str) {
		if (StringUtils.isNotEmpty(str)) {
			str = str.trim().toLowerCase();
		} else return null;
		Color c = null;
		int r = 0, g = 0, b = 0;
		Pattern pattern = Pattern.compile("^rgb\\(\\s*(\\d+)\\s*,\\s*(\\d+)\\s*,\\s*(\\d+)\\)$");			
		Matcher matcher = pattern.matcher(str);
		try {
			if(matcher.find()){ 
				r = new Integer(matcher.group(1));
				g = new Integer(matcher.group(2));
				b = new Integer(matcher.group(3));
			} else if (str.matches("^#[\\da-z]+$")) {
				if (str.length() == 4) {//简写模式 #06c
					r = Integer.parseInt(str.substring(1, 2) + str.substring(1, 2), 16);
					g = Integer.parseInt(str.substring(2, 3) + str.substring(2, 3), 16);
					b = Integer.parseInt(str.substring(3) + str.substring(3), 16);
				} else if (str.length() == 7) {//非简写模式 #0066cc
					r = Integer.parseInt(str.substring(1, 3), 16);
					g = Integer.parseInt(str.substring(3, 5), 16);
					b = Integer.parseInt(str.substring(5), 16);
				}
			}
			c = new Color(r, g, b);
		} catch (NumberFormatException e) {			
			e.printStackTrace();
		}
		
		return c;
	}
	
	public static void main(String[] args) {
		System.out.println(ColorUtil.String2Color("#06C"));
		System.out.println(ColorUtil.String2Color("#009652"));
		System.out.println(ColorUtil.String2Color("rgb(255,10,102)"));
	}
}
