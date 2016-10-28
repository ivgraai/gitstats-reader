package hu.ivgraai.gitstats.reader.excel;

import java.io.IOException;
import java.io.OutputStream;

/**
 * @author igergo
 * @since Sept 15, 2016
 */
public interface IWorkbook {

    enum CellColor { Grey, Undefined }
    enum CellAlignment { Left, Right, Center }

    ISheet createSheet(String sheetName);
    ISheet getSheet(int index);

    void write(OutputStream stream) throws IOException;

}
