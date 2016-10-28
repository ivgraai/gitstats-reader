package hu.ivgraai.gitstats.reader;

import hu.ivgraai.gitstats.reader.excel.ExcelExporter;
import hu.ivgraai.gitstats.reader.excel.ExportRow;
import hu.ivgraai.gitstats.reader.excel.ISheet;
import hu.ivgraai.gitstats.reader.excel.IWorkbook;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author igergo
 * @since Sept 13, 2016
 */
public class Main {

    public static void main(String args[]) throws FileNotFoundException, IOException {
        if (2 != args.length) {
            throw new UnsupportedOperationException();
        }

        IWorkbook workbook = ExcelExporter.createWorkbook();
        exportDatas(workbook, "Commits_per_Author", args[0], "commits_by_author");
        exportDatas(workbook, "Lines_of_Code", args[0], "lines_of_code_by_author");

        File file = new File(args[1]);
        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
            fos.flush();
        }
    }

    public static void exportDatas(IWorkbook workbook, String sheetName, String pathname, String filePrefix) throws FileNotFoundException {
        List<String> lines = new ArrayList<>();
        File plotInput = new File(pathname + filePrefix + ".plot");
        try (Scanner scanner = new Scanner(plotInput)) {
            while (scanner.hasNextLine()) {
                lines.add(scanner.nextLine());
            }
        }

        List<String> authors = new ArrayList<>();

        Pattern pattern = Pattern.compile("\\\".*?\\\"");
        Matcher matcher = pattern.matcher(lines.get(lines.size() - 1));
        while (matcher.find()) {
            authors.add(matcher.group(0).replaceAll("\"", ""));
        }

        ISheet sheet = workbook.createSheet(sheetName);
        List<ExportRow> rows = new ArrayList<>();

        List<ExportRow.ExportCell> header = new ArrayList<>();
        header.add(new ExportRow.ExportCell("", String.class, IWorkbook.CellColor.Grey, IWorkbook.CellAlignment.Center));

        authors.stream().forEach((temp) -> {
            header.add(new ExportRow.ExportCell(temp, String.class, IWorkbook.CellColor.Grey, IWorkbook.CellAlignment.Center));
        });
        rows.add(new ExportRow(header));

        File dataInput = new File(pathname + filePrefix + ".dat");
        int count = authors.size();

        try (Scanner scanner = new Scanner(dataInput)) {

            while (scanner.hasNextLong()) {

                List<ExportRow.ExportCell> row = new ArrayList<>();
                Date date = new Date(scanner.nextLong() * 1000);

                row.add(new ExportRow.ExportCell(date, Date.class, IWorkbook.CellColor.Undefined, IWorkbook.CellAlignment.Center));
                for (int i = 0; i < count; ++i) {
                    int value = scanner.nextInt();
                    row.add(new ExportRow.ExportCell(value, Integer.class, IWorkbook.CellColor.Undefined, IWorkbook.CellAlignment.Center));
                }

                rows.add(new ExportRow(row));

            }

            sheet.setRows(rows);

        }
    }

}
