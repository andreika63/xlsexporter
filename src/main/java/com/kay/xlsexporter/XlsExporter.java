package com.kay.xlsexporter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.*;
import java.util.function.Function;
import java.util.function.Supplier;
import java.util.stream.Stream;

public class XlsExporter<T> {

    private final Class<T> sourceClass;

    private final List<Column<T>> columnList = new ArrayList<>();
    private Supplier<Iterable<T>> dataProvider = List::of;
    private final SXSSFWorkbook workbook = new SXSSFWorkbook(1000);
    private final CellStyle defaultStyle;
    private final CellStyle dateStyle;
    private final CellStyle dateTimeStyle;

    private XlsExporter(Class<T> sourceClass) {
        this.sourceClass = sourceClass;
        defaultStyle = workbook.createCellStyle();
        defaultStyle.setWrapText(true);

        dateStyle = workbook.createCellStyle();
        dateStyle.cloneStyleFrom(defaultStyle);
        dateStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("yyyy-MM-dd"));

        dateTimeStyle = workbook.createCellStyle();
        dateTimeStyle.cloneStyleFrom(defaultStyle);
        dateTimeStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("yyyy-MM-dd hh:mm:ss"));
    }

    public static <T> XlsExporter<T> of(Class<T> sourceClass) {
        return new XlsExporter<>(sourceClass);
    }

    public XlsExporter<T> addColumn(String header, Function<T, ?> valueProvider) {
        columnList.add(new Column<>(header != null ? header : "", valueProvider, 30));
        return this;
    }

    public XlsExporter<T> addColumn(String header, Function<T, ?> valueProvider, Integer width) {
        columnList.add(new Column<>(header != null ? header : "", valueProvider, width));
        return this;
    }

    public XlsExporter<T> addField(ExportField exportField) {
        addColumn(exportField.getCation(), getValueProvider(exportField.getName()));
        return this;
    }

    public XlsExporter<T> addFields(Collection<ExportField> exportFields) {
        exportFields.forEach(this::addField);
        return this;
    }

    public XlsExporter<T> dataProvider(Supplier<Iterable<T>> dataProvider) {
        this.dataProvider = dataProvider;
        return this;
    }

    public XlsExporter<T> dataProvider(Stream<T> stream) {
        this.dataProvider = () -> stream::iterator;
        return this;
    }

    public byte[] export() {
        try {
            createSheet();
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            return outputStream.toByteArray();
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            workbook.dispose();
        }
    }

    public void export(OutputStream outputStream) {
        try(OutputStream out = outputStream) {
            createSheet();
            workbook.write(out);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            workbook.dispose();
        }
    }

    private Function<T, ?> getValueProvider(String name) {

        ArrayDeque<Field> fieldPath = new ArrayDeque<>();

        Stream.of(name.split("\\.")).forEach(fieldName -> {
            if (fieldPath.isEmpty()) {
                fieldPath.addLast(getField(sourceClass, fieldName));
            } else {
                fieldPath.addLast(getField(fieldPath.getLast().getType(), fieldName));
            }
        });
        return t -> {
            Object result = t;
            for (Field field : fieldPath) {
                try {
                    if (result != null) {
                        result = field.get(result);
                    }
                } catch (IllegalAccessException e) {
                    throw new RuntimeException(e);
                }
            }
            return result;
        };
    }

    private Field getField(Class<?> sourceClass,  String fieldName) {
        return Stream.of(sourceClass.getDeclaredFields())
                .filter(df -> df.getName().equals(fieldName))
                .peek(f -> f.setAccessible(true))
                .findFirst()
                .orElseThrow(() -> new IllegalStateException("Class '%s' doesn't have declared field '%s'".formatted(sourceClass.getCanonicalName(), fieldName)));
    }

    private void createSheet() {
        SXSSFSheet sheet = workbook.createSheet();
        sheet.createFreezePane(0, 1);

        int rowNum = 0;
        if (columnList.isEmpty()) {
            getAllProperties(sourceClass, "")
                    .forEach(prop -> addField(new ExportField(prop, prop)));
        }
        fillRow(sheet.createRow(rowNum++), columnList.stream().map(Column::getHeader).toList());

        for (T t : dataProvider.get()) {
            List rowValues = new ArrayList<>();
            for (int i = 0; i< columnList.size() ; i++) {
                rowValues.add(columnList.get(i).getValueProvider().apply(t));
                Integer width = columnList.get(i).getWidth();
                sheet.setColumnWidth(i, (width > 255 ? 255 : width) * 256);
            }
            fillRow(sheet.createRow(rowNum++), rowValues);
        }
    }

    private Stream<String> getAllProperties(Class<?> sourceClass, String prefix) {

        return Stream.of(sourceClass.getDeclaredFields())
                .flatMap(field -> {
                    if (field.getType().getPackage().getName().startsWith("java.")) {
                        return Stream.of(prefix + field.getName());
                    }
                    return getAllProperties(field.getType(), prefix + field.getName() + ".");
                });
    }


    private void fillRow(Row row, List<?> rowValues) {
        int col = 0;
        for (Object it : rowValues) {
            if (it == null) {
                row.createCell(col++, CellType.BLANK);
            } else if (it instanceof LocalDate) {
                addDateCell(row, col++, (LocalDate) it);
            } else if (it instanceof LocalDateTime) {
                addDateTimeCell(row, col++, (LocalDateTime) it);
            } else if (it instanceof ZonedDateTime) {
                addDateTimeCell(row, col++, ((ZonedDateTime) it).toLocalDateTime());
            } else if (it instanceof Date) {
                addDateTimeCell(row, col++, ((Date) it).toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime());
            } else if (it instanceof Number) {
                Cell cell = row.createCell(col++, CellType.NUMERIC);
                cell.setCellStyle(defaultStyle);
                cell.setCellValue(it.toString());
            } else if (it instanceof Boolean) {
                Cell cell = row.createCell(col++, CellType.STRING);
                cell.setCellStyle(defaultStyle);
                cell.setCellValue((Boolean) it ? "да" : "нет");
            } else {
                Cell cell = row.createCell(col++, CellType.STRING);
                cell.setCellStyle(defaultStyle);
                cell.setCellValue(it.toString());
            }
        }
    }

    private void addDateCell(Row row, int col, LocalDate value) {
        Cell cell = row.createCell(col, CellType.NUMERIC);
        cell.setCellStyle(dateStyle);
        cell.setCellValue(value);
    }

    private void addDateTimeCell(Row row, int col, LocalDateTime value) {
        Cell cell = row.createCell(col, CellType.NUMERIC);
        cell.setCellStyle(dateTimeStyle);
        cell.setCellValue(value);
    }


}
