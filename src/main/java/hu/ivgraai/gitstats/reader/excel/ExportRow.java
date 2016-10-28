package hu.ivgraai.gitstats.reader.excel;

import java.util.List;

/**
 * @author igergo
 * @since Sept 15, 2016
 */
public class ExportRow {

    public static class ExportCell {

        public Object value;
        public Class type;
        public IWorkbook.CellColor backgroundColor;
        public IWorkbook.CellAlignment alignment;

        public ExportCell() {
            // empty method
        }

        public ExportCell(Object value, Class type) {
            this.value = value;
            this.type = type;
            this.backgroundColor = IWorkbook.CellColor.Undefined;
            // this.alignment = IWorkbook.CellAlignment.Left;
        }

        public ExportCell(Object value, Class type, IWorkbook.CellColor backgroundColor, IWorkbook.CellAlignment alignment) {
            this.value = value;
            this.type = type;
            this.backgroundColor = backgroundColor;
            this.alignment = alignment;
        }

    }

    private ExportCell[] elements;

    public ExportRow(List<ExportCell> elements) {
        this.elements = elements.toArray(new ExportCell[elements.size()]);
    }

    public Object[] getValues() {
        Object[] result = new Object[elements.length];
        for (int i = 0; i < elements.length; ++i) {
            result[i] = elements[i].value;
        }

        return result;
    }

    public ExportCell[] getElements() {
        return elements;
    }

    public void setElements(ExportCell[] elements) {
        this.elements = elements;
    }

}
