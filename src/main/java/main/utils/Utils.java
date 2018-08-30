/**
 */
package main.utils;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.Normalizer;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Locale;
import java.util.Properties;
import java.util.regex.Pattern;
import org.apache.log4j.Logger;


/**
 * 
 */
public class Utils {

	private static final Logger log = Logger.getLogger(Utils.class);

	public static String encrypt(String str) throws Exception {

		MessageDigest md = MessageDigest.getInstance("MD5");
		md.update(str.getBytes());
		byte byteData[] = md.digest();
		StringBuffer sb = new StringBuffer();
		for (int i = 0; i < byteData.length; i++) {
			sb.append(Integer.toString((byteData[i] & 0xff) + 0x100, 16).substring(1));
		}
		return sb.toString();
	}

	/**
	 * 
	 * @param pdate
	 * @param pformat
	 * @return
	 */
	public static String formatDate(Date pdate, String pformat) {
		SimpleDateFormat sdf = new SimpleDateFormat(pformat);
		return sdf.format(pdate);
	}

	public static void printIbatisMapParams(HashMap hm) {

		Iterator iter = hm.keySet().iterator();
		while (iter.hasNext()) {
			String key = (String) iter.next();
			System.out.println("hm :: " + key + " :: " + hm.get(key));
		}
	}

	/**
	 * Format a double to the given pattern and the symbols
	 * 
	 * @param number
	 * @param str
	 * @return the number formatted
	 */
	public static String formatDecimal(double number, String str) {
		Locale local = new Locale("vi", "VN");
		DecimalFormat df = new DecimalFormat(str, new DecimalFormatSymbols(local));
		return df.format(number);
	}

	/**
	 * Format a double to x,xxx.xx format
	 * 
	 * @param number
	 * @param frac
	 *            the number of decimal places
	 * @return the number in x,xxx.xx format
	 */
	public static String writeDecimal(double number, int frac) {
		if (number == 0.00)
			return "-";
		switch (frac) {
		case 0:
			return formatDecimal(number, "#,###");
		case 1:
			return formatDecimal(number, "#,##0.0");
		case 3:
			return formatDecimal(number, "#,##0.000");
		case 5:
			return formatDecimal(number, "#,##0.00000");
		case 7:
			return formatDecimal(number, "#,###.0");
		case 8:
			return formatDecimal(number, "#,000");
		default:
			return formatDecimal(number, "#,###");
		}
	}

	public static String portalPattern(int frac) {
		String pattern = "#,###";
		switch (frac) {
		case 0:
			pattern = "#,###";
		case 1:
			pattern = "#,##0.0";
		case 3:
			pattern = "#,##0.000";
		case 5:
			pattern = "#,##0.00000";
		case 7:
			pattern = "#,###.0";
		default:
			pattern = "#,###";
		}
		return pattern;
	}

	public static String formatDate(String pdate) {

		String sDate = "";
		if (pdate != null && !pdate.trim().equals("")) {
			String year = pdate.substring(0, 4);
			String month = pdate.substring(5, 7);
			String day = pdate.substring(8, 10);
			sDate = day + "/" + month + "/" + year;
		}

		return sDate;
	}

	public static int calAge(String pdate) {
		int year = Integer.parseInt(pdate.substring(6, 10));
		int month = Integer.parseInt(pdate.substring(3, 5));
		int day = Integer.parseInt(pdate.substring(0, 2));
		Calendar cal = new GregorianCalendar(year, month, day);
		Calendar now = new GregorianCalendar();
		int res = now.get(Calendar.YEAR) - cal.get(Calendar.YEAR);
		if ((cal.get(Calendar.MONTH) > now.get(Calendar.MONTH)) || (cal.get(Calendar.MONTH) == now.get(Calendar.MONTH) && cal.get(Calendar.DAY_OF_MONTH) > now.get(Calendar.DAY_OF_MONTH))) {
			res--;
		}
		return res;
	}

	public static BigDecimal round(BigDecimal val, int scale) {
		BigDecimal ret = new BigDecimal(0);
		ret = val.setScale(-scale, BigDecimal.ROUND_HALF_UP);
		return ret;
	}

	public static Properties loadProperties(String strFileName) {
		// get class loader
		ClassLoader loader = Utils.class.getClassLoader();
		if (loader == null)
			loader = ClassLoader.getSystemClassLoader();
		Properties prop = new Properties();
		java.net.URL url = loader.getResource(strFileName);
		try {
			prop.load(url.openStream());
		} catch (Exception e) {
			e.printStackTrace();
			System.err.println("Could not load configuration file: " + strFileName);
		}
		return prop;
	}

	public static Properties loadSystemProperties(String propFilename) {

		final String cur_dir = System.getProperty("user.dir");
		// System.out.println("xem cur_dir: " + cur_dir);
		final String fullPropFilename = cur_dir + System.getProperty("file.separator") + propFilename;
		Properties prop = new Properties(System.getProperties());
		InputStream is;
		try {
			is = new FileInputStream(fullPropFilename);
			prop.load(is);
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return prop;
	}

	public static boolean validatePhonenumber(String pnum) {
		Properties systemProperties = Utils.loadSystemProperties("system.properties");
		String[] mobile_phone_prefix;
		// String phone_number = null;
		try {
			// check length of phone number
			if (pnum == null || pnum.length() < 10 || pnum.length() > 11) {
				log.error("Phone number " + pnum + " is invalid!");
				return false;
			}
			// check digit character
			char[] cs = pnum.toCharArray();
			int i = 0;
			while (i < cs.length && Character.isDigit(cs[i])) {
				i++;
			}
			if (i < cs.length) {
				log.error("1.Validation phone " + pnum + " fail :: invalid character!!!");
				return false;
			}
			// check phone prefix
			if (!pnum.startsWith("0")) {
				log.error("2.Validation phone " + pnum + " fail :: invalid character!!!");
				return false;
			}

			if (pnum.length() == 10) {
				mobile_phone_prefix = systemProperties.getProperty("mobile_phone_3_digits_prefix").split(",");
				for (String pre : mobile_phone_prefix) {
					if (pre.equalsIgnoreCase(pnum.substring(0, 3))) {
						return true;
					}
				}
				return false;
			}

			if (pnum.length() == 11) {
				mobile_phone_prefix = systemProperties.getProperty("mobile_phone_4_digits_prefix").split(",");
				for (String pre : mobile_phone_prefix) {
					if (pre.equalsIgnoreCase(pnum.substring(0, 4))) {
						return true;
					}
				}
				return false;
			}

			// phone_number = Constant.getVn_sms_prefix() + pnum.substring(1);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return true;
	}

	


	public static List<String> splitSms(StringBuffer sms, int numOfchar4sms, int offset, int smsMaxLen) {
		List<String> smses = new LinkedList<String>();
		if (sms.length() <= smsMaxLen) {
			smses.add(sms.toString().trim());
		} else if (sms.charAt(numOfchar4sms + offset) == ' ') {
			smses.add(sms.substring(0, numOfchar4sms + offset).trim());
			smses.addAll(splitSms(new StringBuffer(sms.substring(numOfchar4sms + offset).trim()), numOfchar4sms, 0, smsMaxLen));
		} else {
			int lastIndexOfSpace = sms.substring(0, numOfchar4sms + offset).lastIndexOf(" ");
			if (lastIndexOfSpace >= 0) {
				offset = lastIndexOfSpace - numOfchar4sms;
			} else {
				offset = 0;
			}
			smses.add(sms.substring(0, numOfchar4sms + offset).trim());
			smses.addAll(splitSms(new StringBuffer(sms.substring(numOfchar4sms + offset).trim()), numOfchar4sms, 0, smsMaxLen));
		}
		return smses;
	}

	public static StringBuffer aggregateSmsSentlog(StringBuffer availabelLog, String newSentlog) {
		return availabelLog.append(newSentlog + "\t\n");
	}

	public static String unAccent(String s) {
		String noz = Normalizer.normalize(s, Normalizer.Form.NFD);
		Pattern pattern = Pattern.compile("\\p{InCombiningDiacriticalMarks}+");
		return pattern.matcher(noz).replaceAll("").replaceAll("Đ", "D").replace("đ", "d");
	}

	public static boolean isJobRunningDate(Date runningdt) {
		try {
			SimpleDateFormat sdf = new SimpleDateFormat("MMyyyy");
			SimpleDateFormat sdf2 = new SimpleDateFormat("yyyy/MM/dd");

			Properties systemProperties = Utils.loadSystemProperties("system.properties");
			log.info("Job running schedule: " + systemProperties.getProperty("ag2us_tr_" + sdf.format(runningdt)));
			String[] runningDates = systemProperties.getProperty("ag2us_tr_" + sdf.format(runningdt)).split(",");
			if (runningDates == null) {
				return false;
			}
			for (String rd : runningDates) {
				if (rd.trim().equalsIgnoreCase(sdf2.format(runningdt))) {
					return true;
				}
			}
			return false;
		} catch (Exception e) {
			log.error(e);
			return false;
		}
	}

	public static void writeLogTbl (StringBuffer sblog, int w, String... logct) {
		StringBuffer line = new StringBuffer("+");
		for (int i=0; i<w; i++) {
			line.append("-");
		}
		line.append("+");
		line.append("\t\n");
		sblog.append(line);
		
		for (String st: logct) {
			StringBuffer cnt = new StringBuffer("| ");
			cnt.append(st); 
			while (cnt.length() - w <=0) {
				cnt.append(" ");
			}
			cnt.append("|");
			cnt.append("\t\n");
			sblog.append(cnt);
		}
		sblog.append(line);
	}
	
	public static void main(String[] agrs) {
		Calendar cal = Calendar.getInstance();
		System.out.println(cal.get(Calendar.MONTH));
		

	}
}
