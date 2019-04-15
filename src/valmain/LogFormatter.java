package valmain;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Formatter;
import java.util.logging.LogRecord;

public class LogFormatter extends Formatter {
	// ANSI escape code
	public static final String ANSI_RESET = "\u001B[0m";
	public static final String ANSI_BLACK = "\u001B[30m";
	public static final String ANSI_RED = "\u001B[31m";
	public static final String ANSI_GREEN = "\u001B[32m";
	public static final String ANSI_YELLOW = "\u001B[33m";
	public static final String ANSI_BLUE = "\u001B[34m";
	public static final String ANSI_PURPLE = "\u001B[35m";
	public static final String ANSI_CYAN = "\u001B[36m";
	public static final String ANSI_WHITE = "\u001B[37m";
	private static final String Bright_Black= "\u001b[30;1m";
	private static final String Bright_Red= "\u001b[31;1m";
	private static final String Bright_Green= "\u001b[32;1m";
	private static final String Bright_Yellow= "\u001b[33;1m";
	private static final String Bright_Blue= "\u001b[34;1m";
	private static final String Bright_Magenta= "\u001b[35;1m";
	private static final String Bright_Cyan= "\u001b[36;1m";
	private static final String Bright_White= "\u001b[37;1m";


	private static final String FORMAT = "\u001b[38;5;34m [%1$tF %1$tT] [%2$-7s] %3$s %n";

	@Override
	public synchronized String format(LogRecord lr) {
		return String.format(FORMAT,
			new Date(lr.getMillis()),
				getColorsForLogger(lr) ,
				ANSI_WHITE+ lr.getMessage() + ANSI_RESET
				);

	}
	
	private String getColorsForLogger (LogRecord log ) {
		String logLevelName = log.getLevel().getLocalizedName();
		
		switch (logLevelName) {
		case "INFO":
			
			return ANSI_CYAN+  logLevelName;
		case "WARNING":
			return ANSI_YELLOW + logLevelName;
		case "SEVERE":

			return ANSI_RED + logLevelName;
		case "FINE":

			return Bright_Green + logLevelName;
		case "CONFIG":

			return ANSI_PURPLE + logLevelName;
		default:
			return ANSI_WHITE + logLevelName;
		}
	}
	
	

}


