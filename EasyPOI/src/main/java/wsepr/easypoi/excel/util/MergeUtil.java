package wsepr.easypoi.excel.util;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.poi.hssf.model.RecordStream;
import org.apache.poi.hssf.model.Sheet;
import org.apache.poi.hssf.model.Workbook;
import org.apache.poi.hssf.record.LabelRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.RecordFactory;
import org.apache.poi.hssf.record.UnicodeString;
import org.apache.poi.hssf.record.aggregates.RecordAggregate.RecordVisitor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class MergeUtil {

	public static void merge(File output, File... inputs) {
		if (inputs.length == 0) {
			throw new IllegalArgumentException("必须提供一个以上的输入文件");
		}

		try {
			List<Record> rootRecords = getRecords(new FileInputStream(inputs[0]));
			Workbook workbook = Workbook.createWorkbook(rootRecords);
			List<Sheet> sheets = getSheets(workbook, rootRecords);			
			createSheet(workbook, sheets);
			write(new FileOutputStream(output), getBytes(workbook, sheets));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	private static void createSheet(Workbook workbook, List<Sheet> sheets){
		Sheet newSheet = Sheet.createSheet();
		newSheet.getWindowTwo().setSelected(false);
		newSheet.getWindowTwo().setActive(false);
		sheets.add(newSheet);
		workbook.setSheetName(sheets.size() -1, "Sheet" + (sheets.size() ));
	}

	private static List<Record> getRecords(InputStream input) {
		try {
			POIFSFileSystem poifs = new POIFSFileSystem(input);
			InputStream stream = poifs.getRoot().createDocumentInputStream("Workbook");
			return RecordFactory.createRecords(stream);
		} catch (IOException e) {
			e.printStackTrace();
		}
		return Collections.emptyList();
	}

	private static List<Sheet> getSheets(Workbook workbook, List<Record> records) {
		int recOffset = workbook.getNumRecords();
		// convert all LabelRecord records to LabelSSTRecord
		convertLabelRecords(records, recOffset, workbook);
		List<Sheet> sheets = new ArrayList<Sheet>();
		RecordStream rs = new RecordStream(records, recOffset);
		while (rs.hasNext()) {
			Sheet sheet = Sheet.createSheet(rs);
			sheets.add(sheet);
		}
		return sheets;
	}

	private static byte[] getBytes(Workbook workbook, List<Sheet> sheets) {
		int nSheets = sheets.size();

		// before getting the workbook size we must tell the sheets that
		// serialization is about to occur.
		for (int i = 0; i < nSheets; i++) {
			sheets.get(i).preSerialize();
		}

		int totalsize = workbook.getSize();

		// pre-calculate all the sheet sizes and set BOF indexes
		SheetRecordCollector[] srCollectors = new SheetRecordCollector[nSheets];
		for (int k = 0; k < nSheets; k++) {
			workbook.setSheetBof(k, totalsize);
			SheetRecordCollector src = new SheetRecordCollector();
			sheets.get(k).visitContainedRecords(src, totalsize);
			totalsize += src.getTotalSize();
			srCollectors[k] = src;
		}

		byte[] retval = new byte[totalsize];
		int pos = workbook.serialize(0, retval);

		for (int k = 0; k < nSheets; k++) {
			SheetRecordCollector src = srCollectors[k];
			int serializedSize = src.serialize(pos, retval);
			if (serializedSize != src.getTotalSize()) {
				// Wrong offset values have been passed in the call to
				// setSheetBof() above.
				// For books with more than one sheet, this discrepancy would
				// cause excel
				// to report errors and loose data while reading the workbook
				throw new IllegalStateException("Actual serialized sheet size (" + serializedSize
						+ ") differs from pre-calculated size (" + src.getTotalSize() + ") for sheet (" + k + ")");
				// TODO - add similar sanity check to ensure that
				// Sheet.serializeIndexRecord() does not write mis-aligned
				// offsets either
			}
			pos += serializedSize;
		}
		return retval;
	}

	public static void write(OutputStream out, byte[] bytes) throws IOException {
		POIFSFileSystem fs = new POIFSFileSystem();
		// Write out the Workbook stream
		try {
			fs.createDocument(new ByteArrayInputStream(bytes), "Workbook");
			fs.writeFilesystem(out);
			out.flush();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	private static void convertLabelRecords(List<Record> records, int offset, Workbook workbook) {
		for (int k = offset; k < records.size(); k++) {
			Record rec = records.get(k);

			if (rec.getSid() == LabelRecord.sid) {
				LabelRecord oldrec = (LabelRecord) rec;

				records.remove(k);
				LabelSSTRecord newrec = new LabelSSTRecord();
				int stringid = workbook.addSSTString(new UnicodeString(oldrec.getValue()));

				newrec.setRow(oldrec.getRow());
				newrec.setColumn(oldrec.getColumn());
				newrec.setXFIndex(oldrec.getXFIndex());
				newrec.setSSTIndex(stringid);
				records.add(k, newrec);
			}
		}
	}

	public static void main(String[] args) {
		File output = new File("E:/merge.xls");
		File input = new File("E:/1.xls");
		merge(output, input);
	}

	/**
	 * Totals the sizes of all sheet records and eventually serializes them
	 */
	private static final class SheetRecordCollector implements RecordVisitor {

		private List<Record> _list;
		private int _totalSize;

		public SheetRecordCollector() {
			_totalSize = 0;
			_list = new ArrayList<Record>(128);
		}

		public int getTotalSize() {
			return _totalSize;
		}

		public void visitRecord(Record r) {
			_list.add(r);
			_totalSize += r.getRecordSize();
		}

		public int serialize(int offset, byte[] data) {
			int result = 0;
			int nRecs = _list.size();
			for (int i = 0; i < nRecs; i++) {
				Record rec = _list.get(i);
				result += rec.serialize(offset + result, data);
			}
			return result;
		}
	}
}
