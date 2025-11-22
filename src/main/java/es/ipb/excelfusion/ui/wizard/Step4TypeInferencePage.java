
package es.ipb.excelfusion.ui.wizard;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.Normalizer;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Combo;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Group;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Table;
import org.eclipse.swt.widgets.TableColumn;
import org.eclipse.swt.widgets.TableItem;

import es.ipb.excelfusion.ui.wizard.Step3StructureValidationPage.SheetValidationResult;


/**
 * Step 4 of the wizard:
 * - Analyzes up to 50 data rows per column across all selected file@sheet.
 * - Infers a type for each column: DATE, CURRENCY, INTEGER, TEXT.
 * - Shows a table so the user can review and override the inferred type.
 */
public class Step4TypeInferencePage implements WizardPage
{

	private static final int MAX_ROWS_PER_SHEET = 50;

	public enum ColumnType
	{
		DATE, CURRENCY, INTEGER, TEXT
	}

	public static class ColumnDefinition
	{
		private final int	 index;		 // 0-based column index
		private final String name;		 // header name or generic A/B/C
		private ColumnType	 type;		 // chosen / inferred type
		private final String sampleValue;

		public ColumnDefinition (int index, String name, ColumnType type, String sampleValue)
		{
			this.index = index;
			this.name = name;
			this.type = type;
			this.sampleValue = sampleValue;
		}

		public int getIndex ()
		{
			return index;
		}

		public String getName ()
		{
			return name;
		}

		public ColumnType getType ()
		{
			return type;
		}

		public void setType (ColumnType type)
		{
			this.type = type;
		}

		public String getSampleValue ()
		{
			return sampleValue;
		}
	}

	private final Step1FileSelectionPage	   step1Page;
	private final Step2PreviewPage			   step2Page;
	private final Step3StructureValidationPage step3Page;

	private Composite						   control;
	private Table							   columnTable;
	private Combo							   typeCombo;
	private Button							   applyTypeButton;

	private boolean							   inferenceDone = false;
	private boolean							   headerDefined = false;

	private final List <ColumnDefinition>	   columns		 = new ArrayList <> ();
	private final Map <Integer, ColumnType>	   typeByIndex	 = new HashMap <> ();

	public Step4TypeInferencePage (Step1FileSelectionPage step1Page, Step2PreviewPage step2Page,
	                               Step3StructureValidationPage step3Page)
	{
		this.step1Page = step1Page;
		this.step2Page = step2Page;
		this.step3Page = step3Page;
	}

	@Override
	public String getTitle ()
	{
		return "Column type inference";
	}

	@Override
	public String getDescription ()
	{
		return "The tool infers column types from sample data. You can review and adjust them here.";
	}

	@Override
	public void createControl (Composite parent)
	{
		control = new Composite (parent, SWT.NONE);
		control.setLayout (new GridLayout (1, false));

		createInfoLabel (control);
		createColumnTable (control);
		createEditorSection (control);
	}

	private void createInfoLabel (Composite parent)
	{
		Label info = new Label (parent, SWT.WRAP);
		info.setLayoutData (new GridData (SWT.FILL, SWT.TOP, true, false));
		info.setText ("Column types are inferred by analyzing up to 50 data rows per column across all selected sheets.\n" +
		              "You can adjust any type if needed before proceeding.");
	}

	private void createColumnTable (Composite parent)
	{
		Group group = new Group (parent, SWT.NONE);
		group.setText ("Inferred columns");
		group.setLayoutData (new GridData (SWT.FILL, SWT.FILL, true, true));
		group.setLayout (new GridLayout (1, false));

		columnTable = new Table (group, SWT.BORDER | SWT.FULL_SELECTION | SWT.SINGLE | SWT.V_SCROLL | SWT.H_SCROLL);
		columnTable.setHeaderVisible (true);
		columnTable.setLinesVisible (true);

		GridData gd = new GridData (SWT.FILL, SWT.FILL, true, true);
		gd.heightHint = 300;
		columnTable.setLayoutData (gd);

		TableColumn colName = new TableColumn (columnTable, SWT.LEFT);
		colName.setText ("Column");
		colName.setWidth (200);

		TableColumn colType = new TableColumn (columnTable, SWT.LEFT);
		colType.setText ("Type");
		colType.setWidth (120);

		TableColumn colSample = new TableColumn (columnTable, SWT.LEFT);
		colSample.setText ("Sample value");
		colSample.setWidth (300);

		columnTable.addSelectionListener (new SelectionAdapter ()
		{
			@Override
			public void widgetSelected (SelectionEvent e)
			{
				onTableSelectionChanged ();
			}
		});
	}

	private void createEditorSection (Composite parent)
	{
		Group group = new Group (parent, SWT.NONE);
		group.setText ("Edit selected column type");
		group.setLayoutData (new GridData (SWT.FILL, SWT.TOP, true, false));
		group.setLayout (new GridLayout (3, false));

		Label typeLabel = new Label (group, SWT.NONE);
		typeLabel.setText ("Type:");

		typeCombo = new Combo (group, SWT.DROP_DOWN | SWT.READ_ONLY);
		typeCombo.setItems (new String[] {ColumnType.DATE.name (), ColumnType.CURRENCY.name (),
		                                  ColumnType.INTEGER.name (), ColumnType.TEXT.name () });
		typeCombo.setLayoutData (new GridData (150, SWT.DEFAULT));

		applyTypeButton = new Button (group, SWT.PUSH);
		applyTypeButton.setText ("Apply to selected column");
		applyTypeButton.setEnabled (false);

		applyTypeButton.addSelectionListener (new SelectionAdapter ()
		{
			@Override
			public void widgetSelected (SelectionEvent e)
			{
				applyTypeToSelectedColumn ();
			}
		});
	}

	@Override
	public Composite getControl ()
	{
		return control;
	}

	@Override
	public void onEnter ()
	{
		if (!inferenceDone)
		{
			runInference ();
		}
	}

	private void runInference ()
	{
		columns.clear ();
		typeByIndex.clear ();
		columnTable.removeAll ();

		Integer headerRow = step2Page.getHeaderRow ();
		Integer dataStartRow = step2Page.getDataStartRow (); // 1-based
		if (dataStartRow == null || dataStartRow <= 0)
		{
			showError ("Invalid configuration",
			           "Data start row is not properly configured. Please go back and fix it.");
			return;
		}

		headerDefined = headerRow != null && headerRow > 0;
		int headerRowIndex = headerDefined? (headerRow - 1) : -1;
		int dataStartIndex = dataStartRow - 1; // 0-based

		List <SheetValidationResult> sheetsToImport = step3Page.getSheetsToImport ();
		if (sheetsToImport.isEmpty ())
		{
			showError ("No sheets to import",
			           "No file/sheet combinations are selected for import. Please go back and review the validation step.");
			return;
		}

		// 1) Determine column names and max column count
		List <String> columnNames;
		int columnCount;

		if (headerDefined)
		{
			ColumnHeaderInfo headerInfo = readHeaderFromReferenceSheet (sheetsToImport, headerRowIndex);
			if (headerInfo == null || headerInfo.headerValues.isEmpty ())
			{
				showError ("Header not found",
				           "The specified header row seems to be empty or invalid. Please go back and review your configuration.");
				return;
			}
			columnNames = headerInfo.headerValues;
			columnCount = columnNames.size ();
		}
		else
		{
			// No header: we will infer max column count from data rows and then generate generic names (A, B, C, ...)
			columnCount = findMaxColumnCount (sheetsToImport, dataStartIndex);
			columnNames = generateGenericColumnNames (columnCount);
		}

		// 2) Collect sample values and decide types per column
		List <List <SampleCell>> samplesByColumn = new ArrayList <> ();
		for (int i = 0; i < columnCount; i++)
		{
			samplesByColumn.add (new ArrayList <> ());
		}

		DataFormatter formatter = new DataFormatter (Locale.getDefault ());

		for (SheetValidationResult svr : sheetsToImport)
		{
			File file = svr.getFile ();
			String sheetName = svr.getSheetName ();

			// Some error entries in Step3 had sheetName "<all sheets>" – skip those for data scanning
			if ("<all sheets>".equals (sheetName))
			{
				continue;
			}

			try (FileInputStream fis = new FileInputStream (file); Workbook workbook = WorkbookFactory.create (fis))
			{

				Sheet sheet = workbook.getSheet (sheetName);
				if (sheet == null)
				{
					continue;
				}

				int lastRow = sheet.getLastRowNum ();
				int maxRowToScan = Math.min (lastRow, dataStartIndex + MAX_ROWS_PER_SHEET - 1);

				for (int r = dataStartIndex; r <= maxRowToScan; r++)
				{
					Row row = sheet.getRow (r);
					if (row == null)
					{
						continue;
					}
					for (int c = 0; c < columnCount; c++)
					{
						Cell cell = row.getCell (c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
						if (cell == null)
						{
							continue;
						}
						String textValue = formatter.formatCellValue (cell);
						if (textValue == null || textValue.trim ().isEmpty ())
						{
							continue;
						}

						ColumnType cellType = detectCellType (cell, textValue);
						SampleCell sample = new SampleCell (cellType, textValue);

						samplesByColumn.get (c).add (sample);
					}
				}

			}
			catch (IOException e)
			{
				// Log as needed later; for inference we just skip problematic file@sheets
				e.printStackTrace ();
			}
			catch (Exception ex)
			{
				ex.printStackTrace ();
			}
		}

		// 3) Decide final type per column and take a sample value
		for (int c = 0; c < columnCount; c++)
		{
			List <SampleCell> samples = samplesByColumn.get (c);
			ColumnType inferredType = inferTypeFromSamples (samples);
			String sampleText = samples.isEmpty ()? "" : samples.get (0).textValue;

			String colName = (c < columnNames.size ())? columnNames.get (c) : ("Col " + (c + 1));
			ColumnDefinition def = new ColumnDefinition (c, colName, inferredType, sampleText);
			columns.add (def);
			typeByIndex.put (c, inferredType);
		}

		// 4) Fill table
		for (ColumnDefinition def : columns)
		{
			TableItem item = new TableItem (columnTable, SWT.NONE);
			item.setText (0, def.getName ());
			item.setText (1, def.getType ().name ());
			item.setText (2, def.getSampleValue ());
		}

		inferenceDone = true;
	}

	private ColumnHeaderInfo readHeaderFromReferenceSheet (List <SheetValidationResult> sheetsToImport,
	                                                       int headerRowIndex)
	{
		DataFormatter formatter = new DataFormatter (Locale.getDefault ());

		for (SheetValidationResult svr : sheetsToImport)
		{
			if (!svr.isMatches ())
			{
				// Prefer a matching sheet as reference
				continue;
			}
			File file = svr.getFile ();
			String sheetName = svr.getSheetName ();
			if ("<all sheets>".equals (sheetName))
			{
				continue;
			}

			try (FileInputStream fis = new FileInputStream (file); Workbook workbook = WorkbookFactory.create (fis))
			{

				Sheet sheet = workbook.getSheet (sheetName);
				if (sheet == null)
				{
					continue;
				}

				Row row = sheet.getRow (headerRowIndex);
				if (row == null)
				{
					continue;
				}

				List <String> headerValues = new ArrayList <> ();
				short lastCellNum = row.getLastCellNum ();
				if (lastCellNum < 0)
				{
					continue;
				}

				for (int c = 0; c < lastCellNum; c++)
				{
					Cell cell = row.getCell (c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
					String value;
					if (cell == null)
					{
						value = "";
					}
					else
					{
						value = formatter.formatCellValue (cell);
					}
					headerValues.add (value);
				}

				if (!headerValues.isEmpty ())
				{
					ColumnHeaderInfo info = new ColumnHeaderInfo ();
					info.headerValues = headerValues;
					info.file = file;
					info.sheetName = sheetName;
					return info;
				}

			}
			catch (IOException e)
			{
				e.printStackTrace ();
			}
			catch (Exception ex)
			{
				ex.printStackTrace ();
			}
		}

		// Fallback: try first sheet even if it is not marked as match
		for (SheetValidationResult svr : sheetsToImport)
		{
			File file = svr.getFile ();
			String sheetName = svr.getSheetName ();
			if ("<all sheets>".equals (sheetName))
			{
				continue;
			}

			try (FileInputStream fis = new FileInputStream (file); Workbook workbook = WorkbookFactory.create (fis))
			{

				Sheet sheet = workbook.getSheet (sheetName);
				if (sheet == null)
				{
					continue;
				}

				Row row = sheet.getRow (headerRowIndex);
				if (row == null)
				{
					continue;
				}

				List <String> headerValues = new ArrayList <> ();
				short lastCellNum = row.getLastCellNum ();
				if (lastCellNum < 0)
				{
					continue;
				}

				for (int c = 0; c < lastCellNum; c++)
				{
					Cell cell = row.getCell (c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
					String value;
					if (cell == null)
					{
						value = "";
					}
					else
					{
						value = formatter.formatCellValue (cell);
					}
					headerValues.add (value);
				}

				if (!headerValues.isEmpty ())
				{
					ColumnHeaderInfo info = new ColumnHeaderInfo ();
					info.headerValues = headerValues;
					info.file = file;
					info.sheetName = sheetName;
					return info;
				}

			}
			catch (IOException e)
			{
				e.printStackTrace ();
			}
			catch (Exception ex)
			{
				ex.printStackTrace ();
			}
		}

		return null;
	}

	private int findMaxColumnCount (List <SheetValidationResult> sheetsToImport, int dataStartIndex)
	{
		int max = 0;
		for (SheetValidationResult svr : sheetsToImport)
		{
			File file = svr.getFile ();
			String sheetName = svr.getSheetName ();
			if ("<all sheets>".equals (sheetName))
			{
				continue;
			}

			try (FileInputStream fis = new FileInputStream (file); Workbook workbook = WorkbookFactory.create (fis))
			{

				Sheet sheet = workbook.getSheet (sheetName);
				if (sheet == null)
				{
					continue;
				}

				int lastRow = sheet.getLastRowNum ();
				int maxRowToScan = Math.min (lastRow, dataStartIndex + MAX_ROWS_PER_SHEET - 1);

				for (int r = dataStartIndex; r <= maxRowToScan; r++)
				{
					Row row = sheet.getRow (r);
					if (row == null)
					{
						continue;
					}
					short lastCellNum = row.getLastCellNum ();
					if (lastCellNum > max)
					{
						max = lastCellNum;
					}
				}

			}
			catch (IOException e)
			{
				e.printStackTrace ();
			}
			catch (Exception ex)
			{
				ex.printStackTrace ();
			}
		}
		return max;
	}

	private List <String> generateGenericColumnNames (int count)
	{
		List <String> names = new ArrayList <> ();
		for (int i = 0; i < count; i++)
		{
			names.add (indexToColumnName (i));
		}
		return names;
	}

	private String indexToColumnName (int index)
	{
		// 0 -> A, 25 -> Z, 26 -> AA, etc.
		StringBuilder sb = new StringBuilder ();
		int x = index;
		while (x >= 0)
		{
			int remainder = x % 26;
			sb.insert (0, (char) ('A' + remainder));
			x = (x / 26) - 1;
		}
		return sb.toString ();
	}

	private static class ColumnHeaderInfo
	{
		File		  file;
		String		  sheetName;
		List <String> headerValues;
	}

	private static class SampleCell
	{
		final ColumnType cellType;
		final String	 textValue;

		SampleCell (ColumnType type, String textValue)
		{
			this.cellType = type;
			this.textValue = textValue;
		}
	}

	private ColumnType detectCellType (Cell cell, String formattedValue)
	{
		CellType ct = cell.getCellType ();

		if (ct == CellType.NUMERIC)
		{
			if (DateUtil.isCellDateFormatted (cell))
			{
				return ColumnType.DATE;
			}
			// numeric but not date
			// check if integer-ish
			double d = cell.getNumericCellValue ();
			if (Math.floor (d) == d)
			{
				// integer number
				return ColumnType.INTEGER;
			}
			else
			{
				// decimal number -> we consider it as CURRENCY candidate by default
				return ColumnType.CURRENCY;
			}
		}

		// For formulas, try evaluated type via formatted string heuristics
		if (ct == CellType.FORMULA)
		{
			// We rely on formattedValue heuristics
			return inferTypeFromFormattedString (formattedValue);
		}

		// Strings, booleans, etc. -> use formatted string
		return inferTypeFromFormattedString (formattedValue);
	}

	private ColumnType inferTypeFromFormattedString (String value)
	{
		if (value == null)
		{
			return ColumnType.TEXT;
		}
		String trimmed = value.trim ();
		if (trimmed.isEmpty ())
		{
			return ColumnType.TEXT;
		}

		// Try integer
		if (trimmed.matches ("^-?\\d+$"))
		{
			return ColumnType.INTEGER;
		}

		// Try currency-like pattern: optional currency symbol, number with optional decimals.
		// We also allow currency symbol either prefix or suffix.
		String normalized = removeAccents (trimmed);
		if (normalized.matches ("^[€$£]?\\s*-?\\d+[.,]?\\d*\\s*[€$£]?$"))
		{
			return ColumnType.CURRENCY;
		}

		// Date-like detection is tricky via text; we leave it as TEXT in this branch.
		// (Pure date cells are already handled via NUMERIC + DateUtil above.)
		return ColumnType.TEXT;
	}

	private ColumnType inferTypeFromSamples (List <SampleCell> samples)
	{
		if (samples == null || samples.isEmpty ())
		{
			// default to TEXT if we have no data
			return ColumnType.TEXT;
		}

		boolean anyDate = false;
		boolean anyNonNumericText = false;
		boolean allInteger = true;
		boolean allNumericOrCurrency = true;

		for (SampleCell s : samples)
		{
			ColumnType ct = s.cellType;
			if (ct == ColumnType.DATE)
			{
				anyDate = true;
			}
			else if (ct == ColumnType.INTEGER)
			{
				// integer is numeric, keep flags
			}
			else if (ct == ColumnType.CURRENCY)
			{
				allInteger = false; // currency may be decimal
			}
			else if (ct == ColumnType.TEXT)
			{
				anyNonNumericText = true;
				allNumericOrCurrency = false;
				allInteger = false;
			}
		}

		if (anyDate)
		{
			return ColumnType.DATE;
		}
		if (!anyNonNumericText && allInteger)
		{
			return ColumnType.INTEGER;
		}
		if (!anyNonNumericText && allNumericOrCurrency)
		{
			return ColumnType.CURRENCY;
		}
		return ColumnType.TEXT;
	}

	private String removeAccents (String input)
	{
		String normalized = Normalizer.normalize (input, Normalizer.Form.NFD);
		Pattern pattern = Pattern.compile ("\\p{M}+");
		return pattern.matcher (normalized).replaceAll ("");
	}

	private void onTableSelectionChanged ()
	{
		int index = columnTable.getSelectionIndex ();
		if (index < 0 || index >= columns.size ())
		{
			applyTypeButton.setEnabled (false);
			return;
		}

		ColumnDefinition def = columns.get (index);
		ColumnType type = def.getType ();
		if (type != null)
		{
			String typeName = type.name ();
			int comboIndex = typeCombo.indexOf (typeName);
			if (comboIndex >= 0)
			{
				typeCombo.select (comboIndex);
			}
			else
			{
				typeCombo.deselectAll ();
			}
		}
		else
		{
			typeCombo.deselectAll ();
		}

		applyTypeButton.setEnabled (true);
	}

	private void applyTypeToSelectedColumn ()
	{
		int tableIndex = columnTable.getSelectionIndex ();
		if (tableIndex < 0 || tableIndex >= columns.size ())
		{
			return;
		}

		int comboIndex = typeCombo.getSelectionIndex ();
		if (comboIndex < 0)
		{
			return;
		}

		String selectedTypeName = typeCombo.getItem (comboIndex);
		ColumnType selectedType = ColumnType.valueOf (selectedTypeName);

		ColumnDefinition def = columns.get (tableIndex);
		def.setType (selectedType);
		typeByIndex.put (def.getIndex (), selectedType);

		TableItem item = columnTable.getItem (tableIndex);
		item.setText (1, selectedType.name ());
	}

	@Override
	public boolean onLeave ()
	{
		// Nothing special to validate here; user is allowed to keep inferred types as-is.
		return true;
	}

	@Override
	public boolean canGoNext ()
	{
		// Always allow Next from this page (types will default if user doesn't touch anything).
		return true;
	}

	@Override
	public boolean canGoBack ()
	{
		return true;
	}

	private void showError (String title, String message)
	{
		MessageBox mb = new MessageBox (control.getShell (), SWT.ICON_ERROR | SWT.OK);
		mb.setText (title);
		mb.setMessage (message);
		mb.open ();
	}

	// === Public getters for later steps (DB creation / import) ===

	public List <ColumnDefinition> getColumns ()
	{
		return new ArrayList <> (columns);
	}

	public Map <Integer, ColumnType> getColumnTypesByIndex ()
	{
		return new HashMap <> (typeByIndex);
	}

	public boolean isHeaderDefined ()
	{
		return headerDefined;
	}
}
