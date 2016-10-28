package hu.ivgraai.gitstats.reader.excel;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * @author igergo
 * @since Sept 15, 2016
 */
class PoiWorkbook implements IWorkbook {

    private final HSSFWorkbook workbook;
    private final CellSytleManager manager = new CellSytleManager();
    private final Object workbookLock = new Object();

    public PoiWorkbook() {
        workbook = new HSSFWorkbook();
    }

    public PoiWorkbook(InputStream stream) throws IOException {
        workbook = new HSSFWorkbook(stream);
    }

    @Override
    public void write(OutputStream stream) throws IOException {
        workbook.write(stream);
    }

    @Override
    public ISheet createSheet(String sheetName) {
        return new PoiSheet(workbook.createSheet(sheetName), manager);
    }

    @Override
    public ISheet getSheet(int index) {
        try {
            return new PoiSheet(workbook.getSheetAt(index), manager);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public class CellSytleManager {

        private class CellStyleElement {

            private final short color;
            private final short alignment;

            public CellStyleElement(short color, short alignment) {
                this.color = color;
                this.alignment = alignment;
            }

            @Override
            public boolean equals(Object obj) {
                if (this == obj) {
                    return true;
                }
                if (obj == null) {
                    return false;
                }
                if (getClass() != obj.getClass()) {
                    return false;
                }
                final CellStyleElement other = (CellStyleElement) obj;
                return (
                        this.color == other.color &&
                        this.alignment == other.alignment
                    );
            }

            @Override
            public int hashCode() {
                int hash = 7;
                hash = 79 * hash + this.color;
                hash = 79 * hash + this.alignment;
                return hash;
            }

        }

        private final Map<CellStyleElement, CellStyle> cache = new ConcurrentHashMap<>();

        public CellSytleManager() {
            // empty method
        }

        public CellStyle getStyle(short color, short alignment) {
            CellStyle style = cache.get(new CellStyleElement(color, alignment));
            if (null != style) {
                return style;
            }

            synchronized (workbookLock) {
                style = workbook.createCellStyle();
            }

            style.setAlignment(alignment);
            style.setFillForegroundColor(color);
            style.setFillPattern(CellStyle.SOLID_FOREGROUND);

            style.setBorderLeft((short) 1);
            style.setBorderRight((short) 1);
            style.setBorderTop((short) 1);
            style.setBorderBottom((short) 1);

            cache.put(new CellStyleElement(color, alignment), style);
            return style;
        }

    }

}
