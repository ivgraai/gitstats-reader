package hu.ivgraai.gitstats.reader.excel;

import java.io.IOException;
import java.io.InputStream;
import java.io.PipedInputStream;
import java.io.PipedOutputStream;

/**
 * @author igergo
 * @since Sept 15, 2016
 */
public class ExcelExporter {

    public static InputStream export(final IWorkbook workbook, String prefix, String suffix) {
        PipedInputStream is = new PipedInputStream();
        try {
            final PipedOutputStream outputStream = new PipedOutputStream(is);
            new Thread(() -> {

                try {
                    workbook.write(outputStream);
                } catch (IOException e) {
                    throw new RuntimeException(e);
                } finally {
                    try {
                        outputStream.close();
                    } catch (IOException e) {
                        // empty block
                    }
                }

            }).start();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        return is;
    }

    public static IWorkbook createWorkbook() {
        return new PoiWorkbook();
    }

    public static IWorkbook createWorkbook(InputStream stream) throws IOException {
        return new PoiWorkbook(stream);
    }

}
