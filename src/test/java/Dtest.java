import java.text.*;
import java.util.*;

public class Dtest {
	public static void main(String args[]) {
		SimpleDateFormat df = (SimpleDateFormat) DateFormat.getDateInstance(DateFormat.SHORT);
		System.out.println("The short date format is  " + df.toPattern());
		Locale loc = Locale.getAvailableLocales()[0];
		df = (SimpleDateFormat) DateFormat.getDateInstance(DateFormat.SHORT, loc);
		System.out.println("The short date format is  " + df.toPattern());
		System.out.println(df.format(new Date()));
	}
}