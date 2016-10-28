package hu.ivgraai.gitstats.reader.excel;

import java.util.List;

/**
 * @author igergo
 * @since Sept 15, 2016
 */
public interface ISheet {

    void setRows(List<ExportRow> rows);
    List<Object[]> getRows();

}
