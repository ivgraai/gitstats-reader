package hu.ivgraai.gitstats.reader;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author igergo
 * @since Sept 17, 2016
 */
public class Utility {

    public static String convertToString(Date date, String format) {
        DateFormat df = new SimpleDateFormat(format);
        return df.format(date);
    }

}
