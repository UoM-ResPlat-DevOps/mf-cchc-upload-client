import arc.xml.XmlStringWriter;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

class Spreadsheet implements Iterable<Spreadsheet.Record> {

    private List<Record> _records;

//    rowStartIndex is 0-based
    public Spreadsheet(FileInputStream inp, int rowStartIndex, String dirPrefix, String doctype) throws  Throwable{
        _records = new ArrayList<Record>();
        Workbook wb = WorkbookFactory.create(inp);
        Sheet sheet = wb.getSheetAt(0);
        for (Row r: sheet){
            if (processRow(r, rowStartIndex)){
               _records.add(Record.create(r, dirPrefix, doctype));
            }
        }
    }

    private static boolean processRow(Row r, int startRowIndex) {
        if (r!=null && r.getRowNum() >= startRowIndex){
            return true;
        }
        return false;
    }

    public Iterator<Record> iterator() {
        return _records.iterator();
    }

    public Record getRecordAt(int index) {
        return _records.get(index);
    }

    public int length() {
        return _records.size();
    }


    public static class Record {

        public final List<String> slideImages;
        public final String title;
        public final String subject;
        public final String description;
        public final String location;
        public final String coordinates;
        public final String creator;
        public final String publisher;
        public final String date;
        public final String contributor;
        public final String rights;
        public final String format;
        public final String type;
        public final String identifier;
        public final String filepath;
        public final String mediafluxFolder;
        public final String frontViewFilenameTif;
        public final String backViewFilenameTif;
        public final String frontViewFilenameJpg;
        public final String backViewFilenameJpg;


        public static Record create (Row row, String dirPrefix, String doctype) throws Throwable{
            return new Record (row, dirPrefix, doctype);
        }

        public Record (Row row, String dirPrefix, String doctype) throws Throwable {
            title = readString(row, 0, "title");
            subject = readString(row, 1, "subject");
            description = readString(row, 2, "description");
            location = readString(row, 3, "location");
            coordinates = readString(row, 4, "coordinates");
            creator = readString(row, 5, "creator");
            publisher = readString(row, 6, "publisher");
            date = readString(row, 7, "date");
            contributor = readString(row, 8, "contributor");
            rights = readString(row, 9, "rights");
            format = readString(row, 10, "format");
            type = readString(row, 11, "type");
            identifier = readString(row, 12, "identifier");
            filepath = readString(row, 13, "filepath");

            mediafluxFolder = readMediafluxFolder(row, 14, "mediaflux folder");
            frontViewFilenameTif = readString(row, 15, "front view filename tif");
            backViewFilenameTif = readString(row, 16, "back view filename tifj");
            frontViewFilenameJpg = readString(row, 17, "front view filename jpg");
            backViewFilenameJpg = readString(row, 18, "back view filename jpg");

            slideImages = new ArrayList<String>();
            if (StringUtils.isNotBlank(frontViewFilenameTif)) slideImages.add(frontViewFilenameTif);
            if (StringUtils.isNotBlank(backViewFilenameTif)) slideImages.add(backViewFilenameTif);
            if (StringUtils.isNotBlank(frontViewFilenameJpg)) slideImages.add(frontViewFilenameJpg);
            if (StringUtils.isNotBlank(frontViewFilenameJpg)) slideImages.add(backViewFilenameJpg);
        }


        private String readString(Row row, int colIndex, String columnName) throws Throwable {
            Cell c = row.getCell(colIndex);
            if (c == null ){
                CellReference cr = new CellReference(row.getRowNum(), c.getColumnIndex());
                throw new Exception ("Null pointer exception for cell object for cell: " + cr.formatAsString());
            }
            String val = c.toString().trim();
            return val;
        }

        private String readMediafluxFolder(Row row, int colIndex, String columnName) throws Throwable {
            Cell c = row.getCell(colIndex);
            if (c == null ){
                CellReference cr = new CellReference(row.getRowNum(), c.getColumnIndex());
                throw new Exception ("Null pointer exception for cell object for cell: " + cr.formatAsString());
            }
            String val;
            if (isNumeric(c.toString())){
                val = Integer.toString( (int) c.getNumericCellValue());
            }
            else{
                val = c.toString().trim();
            }
            System.out.println(val);
            return val;
        }


        public XmlStringWriter toXmlStringWriter(String ns,
                                                 String docTypeName,
                                                 String assetName) throws Throwable{
            XmlStringWriter w = new XmlStringWriter();
            w.add("namespace", ns);
            w.add("name", assetName);
            w.push("meta");
            w.push(docTypeName);
            w.add("title", title);
            if (subject.length() > 0) w.add("subject", subject);
            if (description.length() > 0) w.add("description", description);
            if (location.length() > 0) w.add("location", location);
            if (coordinates.length() > 0) w.add("coordinates", coordinates);
            if (creator.length() > 0) w.add("creator", creator);
            if (publisher.length() > 0) w.add("publisher", publisher);
            if (date.length() > 0) w.add("date", date);
            if (contributor.length() > 0) w.add("contributor", contributor);
            if (rights.length() > 0) w.add("rights", rights);
            if (format.length() > 0) w.add("format", format);
            if (type.length() > 0) w.add("type", type);
            if (filepath.length() > 0) w.add("filepath", filepath);
            if (mediafluxFolder.length() > 0) w.add("mediaflux-folder", mediafluxFolder);
            if (frontViewFilenameTif.length() > 0) w.add("front-view-filename-tif", frontViewFilenameTif);
            if (backViewFilenameTif.length() > 0) w.add("back-view-filename-tif", backViewFilenameTif);
            if (frontViewFilenameJpg.length() > 0) w.add("front-view-filename-jpg", frontViewFilenameJpg);
            if (backViewFilenameJpg.length() > 0) w.add("back-view-filename-jpg", backViewFilenameJpg);
            w.pop();
            return w;
        }

        public static boolean isNumeric(String str)
        {
            try
            {
                double d = Double.parseDouble(str);
            }
            catch(NumberFormatException nfe)
            {
                return false;
            }
            return true;
        }
    }
}
