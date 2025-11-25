
package es.ipb.excelfusion.ui.wizard;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.ModifyEvent;
import org.eclipse.swt.events.ModifyListener;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.graphics.Color;
import org.eclipse.swt.graphics.Point;
import org.eclipse.swt.graphics.Rectangle;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.layout.RowLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Combo;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Group;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Table;
import org.eclipse.swt.widgets.TableColumn;
import org.eclipse.swt.widgets.TableItem;
import org.eclipse.swt.widgets.Text;

import es.ipb.excelfusion.config.ImportConfiguration;


/**
 * Step 2 of the wizard:
 * - Preview the selected Excel file (first selected from step 1)
 * - Choose header row and data start row
 * - Configure auto-increment / fill-empty options
 * - Set target table name
 * - Highlight the header row (if defined) in the preview
 */
public class Step2PreviewPage implements WizardPage
{

	private static final int						 MAX_PREVIEW_ROWS	= 500;

	private Composite								 control;

	private final ImportConfiguration				 config;

	private Combo									 sheetCombo;
	private Label									 sheetInfoLabel;

	private Table									 previewTable;

	private Text									 headerRowText;
	private Text									 dataStartRowText;

	private Button									 autoIncrementCheckbox;

	private Text									 tableNameText;

	private File									 selectedFile;
	private final Map <String, List <List <String>>> previewDataBySheet	= new LinkedHashMap <> ();
	private final List <String>						 sheetNames			= new ArrayList <> ();

	private Color									 highlightColor;
	private Color									 nonDataColor;

	private boolean[]								 fillEmptyPerColumn;

	public Step2PreviewPage (ImportConfiguration config)
	{
		this.config = config;
	}

	@Override
	public String getTitle ()
	{
		return "Preview and row configuration";
	}

	@Override
	public String getDescription ()
	{
		return "Preview the Excel structure, select header and data rows, and configure basic options.";
	}

	@Override
	public void createControl (Composite parent)
	{
		control = new Composite (parent, SWT.NONE);
		control.setLayout (new GridLayout (1, false));

		highlightColor = parent.getDisplay ().getSystemColor (SWT.COLOR_YELLOW);
		nonDataColor = parent.getDisplay ().getSystemColor (SWT.COLOR_WIDGET_LIGHT_SHADOW);

		createTopSection (control);
		createPreviewTable (control);
		createOptionsSection (control);
		createTableNameSection (control);
	}

	private void createTopSection (Composite parent)
	{
		Composite top = new Composite (parent, SWT.NONE);
		top.setLayoutData (new GridData (SWT.FILL, SWT.TOP, true, false));
		GridLayout gl = new GridLayout (4, false);
		gl.marginWidth = 0;
		top.setLayout (gl);

		// File info
		Label fileLabel = new Label (top, SWT.NONE);
		fileLabel.setText ("File:");

		Label fileNameLabel = new Label (top, SWT.NONE);
		fileNameLabel.setLayoutData (new GridData (SWT.FILL, SWT.CENTER, true, false));
		fileNameLabel.setText ("(no file selected yet)");

		// Sheet combo
		sheetCombo = new Combo (top, SWT.DROP_DOWN | SWT.READ_ONLY);
		sheetCombo.setLayoutData (new GridData (SWT.LEFT, SWT.CENTER, false, false));
		sheetCombo.addSelectionListener (new SelectionAdapter ()
		{
			@Override
			public void widgetSelected (SelectionEvent e)
			{
				onSheetSelected ();
			}
		});

		sheetInfoLabel = new Label (top, SWT.NONE);
		sheetInfoLabel.setLayoutData (new GridData (SWT.FILL, SWT.CENTER, true, false));
		sheetInfoLabel.setText ("");

		// Save reference to file name label for later
		control.setData ("fileNameLabel", fileNameLabel);

		// Second row: header & data start
		Composite rowConfig = new Composite (parent, SWT.NONE);
		rowConfig.setLayoutData (new GridData (SWT.FILL, SWT.TOP, true, false));
		GridLayout gl2 = new GridLayout (4, false);
		gl2.marginWidth = 0;
		gl2.horizontalSpacing = 10;
		rowConfig.setLayout (gl2);

		Label headerLabel = new Label (rowConfig, SWT.NONE);
		headerLabel.setText ("Header row:");

		headerRowText = new Text (rowConfig, SWT.BORDER);
		headerRowText.setLayoutData (new GridData (60, SWT.DEFAULT));
		headerRowText.setToolTipText ("Row number containing the header (0 or empty = no header)");

		Label dataStartLabel = new Label (rowConfig, SWT.NONE);
		dataStartLabel.setText ("Data starts at row:");

		dataStartRowText = new Text (rowConfig, SWT.BORDER);
		dataStartRowText.setLayoutData (new GridData (60, SWT.DEFAULT));

		// highlight header row when headerRowText changes
		headerRowText.addModifyListener (new ModifyListener ()
		{
			@Override
			public void modifyText (ModifyEvent e)
			{
				applyHeaderHighlight ();
			}
		});
	}

	private void createPreviewTable (Composite parent)
	{
		Group group = new Group (parent, SWT.NONE);
		group.setText ("Preview (first " + MAX_PREVIEW_ROWS + " rows)");
		group.setLayoutData (new GridData (SWT.FILL, SWT.FILL, true, true));
		group.setLayout (new GridLayout (1, false));

		previewTable = new Table (group, SWT.BORDER | SWT.FULL_SELECTION | SWT.V_SCROLL | SWT.H_SCROLL);
		previewTable.setHeaderVisible (true);
		previewTable.setLinesVisible (true);
		previewTable.addListener (SWT.MouseDown, event -> {
			Point pt = new Point (event.x, event.y);
			TableItem item = previewTable.getItem (pt);
			if (item == null)
			{
				return;
			}

			int rowIndex = previewTable.indexOf (item);
			// Solo la fila 0 es la "prefila" de checkboxes
			if (rowIndex != 0)
			{
				return;
			}

			int columnCount = previewTable.getColumnCount ();
			for (int c = 1; c < columnCount; c++) // desde la 1: la 0 es "#"
			{
				Rectangle rect = item.getBounds (c);
				if (rect.contains (pt))
				{
					toggleFillEmptyForColumn (c - 1); // columnas de datos empiezan en 0
					break;
				}
			}
		});

		GridData gd = new GridData (SWT.FILL, SWT.FILL, true, true);
		gd.heightHint = 350;
		previewTable.setLayoutData (gd);
	}

	private void createOptionsSection (Composite parent)
	{
		Group group = new Group (parent, SWT.NONE);
		group.setText ("Options");
		group.setLayoutData (new GridData (SWT.FILL, SWT.TOP, true, false));
		group.setLayout (new RowLayout (SWT.HORIZONTAL));

		autoIncrementCheckbox = new Button (group, SWT.CHECK);
		autoIncrementCheckbox.setText ("Add auto-increment column (ID)");
		autoIncrementCheckbox.setSelection (true);

	}

	private void createTableNameSection (Composite parent)
	{
		Group group = new Group (parent, SWT.NONE);
		group.setText ("Target table");
		group.setLayoutData (new GridData (SWT.FILL, SWT.TOP, true, false));
		group.setLayout (new GridLayout (2, false));

		Label label = new Label (group, SWT.NONE);
		label.setText ("Table name:");

		tableNameText = new Text (group, SWT.BORDER);
		tableNameText.setLayoutData (new GridData (SWT.FILL, SWT.CENTER, true, false));
	}

	@Override
	public Composite getControl ()
	{
		return control;
	}

	@Override
	public void onEnter ()
	{
		initializeFromStep1 (); // Always reload the file

		// Rellenar nombre de tabla desde config si existiera
		if (tableNameText != null && !tableNameText.isDisposed ())
		{
			String tableName = config.getTableName ();
			if (tableName != null && !tableName.isBlank ())
			{
				tableNameText.setText (tableName);
			}
		}

		// Rellenar header/data row si ya estaban
		if (headerRowText != null && !headerRowText.isDisposed ())
		{
			Integer hr = config.getHeaderRow ();
			headerRowText.setText (hr != null && hr > 0? String.valueOf (hr) : "");
		}
		if (dataStartRowText != null && !dataStartRowText.isDisposed ())
		{
			Integer ds = config.getDataStartRow ();
			dataStartRowText.setText (ds != null && ds > 0? String.valueOf (ds) : "");
		}

		if (autoIncrementCheckbox != null && !autoIncrementCheckbox.isDisposed ())
		{
			autoIncrementCheckbox.setSelection (config.isAutoIncrement ());
		}

		applyHeaderHighlight ();
	}

	private void initializeFromStep1 ()
	{
		List <File> selectedFiles = config.getSelectedFiles ();
		if (selectedFiles.isEmpty ())
		{
			showWarning ("No files selected", "Please go back and select at least one Excel file.");
			return;
		}

		selectedFile = selectedFiles.get (0); // first file as reference
		Label fileNameLabel = (Label) control.getData ("fileNameLabel");
		if (fileNameLabel != null && !fileNameLabel.isDisposed ())
		{
			fileNameLabel.setText (selectedFile.getName ());
		}

		loadPreviewData (selectedFile);
		populateSheetCombo ();
		if (!sheetNames.isEmpty ())
		{
			sheetCombo.select (0);
			onSheetSelected ();
		}
	}

	private void loadPreviewData (File file)
	{
		previewDataBySheet.clear ();
		sheetNames.clear ();

		try (FileInputStream fis = new FileInputStream (file); Workbook workbook = WorkbookFactory.create (fis))
		{

			int numberOfSheets = workbook.getNumberOfSheets ();
			for (int i = 0; i < numberOfSheets; i++)
			{
				Sheet sheet = workbook.getSheetAt (i);
				String sheetName = sheet.getSheetName ();
				sheetNames.add (sheetName);

				List <List <String>> rows = new ArrayList <> ();

				int maxRow = Math.min (sheet.getLastRowNum (), MAX_PREVIEW_ROWS - 1);
				DataFormatter formatter = new DataFormatter (Locale.getDefault ());

				for (int r = 0; r <= maxRow; r++)
				{
					Row row = sheet.getRow (r);
					if (row == null)
					{
						rows.add (new ArrayList <> ());
						continue;
					}

					int lastCellNum = row.getLastCellNum ();
					if (lastCellNum < 0)
					{
						rows.add (new ArrayList <> ());
						continue;
					}

					List <String> values = new ArrayList <> ();
					for (int c = 0; c < lastCellNum; c++)
					{
						Cell cell = row.getCell (c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
						if (cell == null)
						{
							values.add ("");
						}
						else
						{
							values.add (formatter.formatCellValue (cell));
						}
					}
					rows.add (values);
				}

				previewDataBySheet.put (sheetName, rows);
			}

		}
		catch (IOException e)
		{
			e.printStackTrace ();
			showError ("Error reading Excel file",
			           "Could not read file:\n" + file.getAbsolutePath () + "\n\n" + e.getMessage ());
		}
		catch (Exception e)
		{
			e.printStackTrace ();
			showError ("Error reading Excel file",
			           "Unexpected error while reading file:\n" + file.getAbsolutePath () + "\n\n" + e.getMessage ());
		}
	}

	private void populateSheetCombo ()
	{
		sheetCombo.removeAll ();
		if (sheetNames.isEmpty ())
		{
			sheetInfoLabel.setText ("No sheets found.");
			return;
		}

		sheetCombo.add ("All sheets");
		for (String name : sheetNames)
		{
			sheetCombo.add (name);
		}

		sheetInfoLabel.setText ("Configuration applies to all selected sheets.");
	}

	private void onSheetSelected ()
	{
		int index = sheetCombo.getSelectionIndex ();
		if (index < 0)
		{
			return;
		}

		String selectedName;
		if (index == 0)
		{
			// "All sheets" -> preview uses first sheet
			if (sheetNames.isEmpty ())
			{
				return;
			}
			selectedName = sheetNames.get (0);
		}
		else
		{
			selectedName = sheetCombo.getItem (index);
		}

		List <List <String>> rows = previewDataBySheet.getOrDefault (selectedName, new ArrayList <> ());
		refreshPreviewTable (rows);
		applyHeaderHighlight ();
	}

	private void refreshPreviewTable (List <List <String>> rows)
	{
		previewTable.removeAll ();
		for (TableColumn col : previewTable.getColumns ())
		{
			col.dispose ();
		}

		// Determinar número máximo de columnas de datos
		int maxColumns = 0;
		for (List <String> row : rows)
		{
			maxColumns = Math.max (maxColumns, row.size ());
		}

		// Columna 0: número de fila
		TableColumn rowNumCol = new TableColumn (previewTable, SWT.RIGHT);
		rowNumCol.setText ("#");
		rowNumCol.setWidth (50);

		// Crear columnas A, B, C, ...
		for (int c = 0; c < maxColumns; c++)
		{
			TableColumn column = new TableColumn (previewTable, SWT.LEFT);
			column.setText (indexToColumnName (c));
			column.setWidth (120);
		}

		// Fondo de cabecera (no es parte de los datos)
		if (nonDataColor != null)
		{
			previewTable.setHeaderBackground (nonDataColor);
		}

		// Inicializar configuración por columnas para "fill empty"
		if (maxColumns > 0)
		{
			if (fillEmptyPerColumn == null || fillEmptyPerColumn.length != maxColumns)
			{
				fillEmptyPerColumn = new boolean[maxColumns];
				for (int i = 0; i < maxColumns; i++)
				{
					fillEmptyPerColumn[i] = false; // por defecto, todas desmarcadas
				}
			}

			// === PREFILA CON CHECKBOXES ===
			TableItem prefItem = new TableItem (previewTable, SWT.NONE);
			prefItem.setText (0, ""); // primera columna vacía o "Cfg"

			for (int c = 0; c < maxColumns; c++)
			{
				prefItem.setText (c + 1, fillEmptyPerColumn[c]? "☑" : "☐");
			}

			// prefila: toda en gris
			if (nonDataColor != null)
			{
				for (int c = 0; c < previewTable.getColumnCount (); c++)
				{
					prefItem.setBackground (c, nonDataColor);
				}
			}
		}
		else
		{
			fillEmptyPerColumn = null;
		}

		// === RELLENAR FILAS DE DATOS ===
		for (int r = 0; r < rows.size (); r++)
		{
			List <String> rowData = rows.get (r);
			TableItem item = new TableItem (previewTable, SWT.NONE);

			// Columna 0 → número de fila Excel (1-based)
			item.setText (0, String.valueOf (r + 1));

			for (int c = 0; c < rowData.size (); c++)
			{
				item.setText (c + 1, rowData.get (c)); // +1 por columna "#"
			}
		}

		// Primera columna en gris (no forma parte de los datos)
		if (nonDataColor != null)
		{
			for (TableItem item : previewTable.getItems ())
			{
				item.setBackground (0, nonDataColor);
			}
		}

		previewTable.layout ();
	}

	private void applyHeaderHighlight ()
	{
		if (previewTable == null || previewTable.isDisposed ())
		{
			return;
		}

		int itemCount = previewTable.getItemCount ();
		int columnCount = previewTable.getColumnCount ();

		if (itemCount == 0 || columnCount == 0)
		{
			return;
		}

		// Resetear colores:
		// - Fila 0 (prefila): todas las columnas en gris
		// - Resto de filas: col 0 gris, resto sin color
		for (int i = 0; i < itemCount; i++)
		{
			TableItem item = previewTable.getItem (i);
			if (i == 0)
			{
				// prefila
				for (int c = 0; c < columnCount; c++)
				{
					item.setBackground (c, nonDataColor);
				}
			}
			else
			{
				// datos
				item.setBackground (0, nonDataColor);
				for (int c = 1; c < columnCount; c++)
				{
					item.setBackground (c, null);
				}
			}
		}

		// Calcular índice de la fila de cabecera en la tabla:
		// Excel 1-based → índice de datos 0-based → en la tabla se desplaza +1 por la prefila
		int headerIndexExcelZeroBased = getHeaderRowIndexInternal ();
		if (headerIndexExcelZeroBased < 0)
		{
			return;
		}

		int tableRowIndex = headerIndexExcelZeroBased + 1; // +1 por la prefila
		if (tableRowIndex < 1 || tableRowIndex >= itemCount)
		{
			return;
		}

		// Pintamos la cabecera de datos en amarillo, SOLO columnas de datos (>=1)
		TableItem headerItem = previewTable.getItem (tableRowIndex);
		for (int c = 1; c < columnCount; c++)
		{
			headerItem.setBackground (c, highlightColor);
		}
	}

	private int getHeaderRowIndexInternal ()
	{
		String txt = headerRowText.getText ().trim ();
		if (txt.isEmpty ())
		{
			return -1; // no header
		}
		try
		{
			int rowNumber = Integer.parseInt (txt);
			if (rowNumber <= 0)
			{
				return -1;
			}
			// Excel-style 1-based → table index 0-based
			return rowNumber - 1;
		}
		catch (NumberFormatException e)
		{
			return -1;
		}
	}

	private Integer parseOptionalRowNumber (String text)
	{
		if (text == null || text.isEmpty ())
		{
			return null;
		}
		try
		{
			int value = Integer.parseInt (text);
			if (value < 0)
			{
				showError ("Invalid header row", "Header row must be 0 or a positive integer.");
				return null;
			}
			return value;
		}
		catch (NumberFormatException e)
		{
			showError ("Invalid header row", "Header row must be a valid integer (0 = no header).");
			return null;
		}
	}

	private Integer parseRequiredRowNumber (String text, String label)
	{
		if (text == null || text.isEmpty ())
		{
			showError ("Invalid " + label, label + " is required and must be a positive integer.");
			return null;
		}
		try
		{
			int value = Integer.parseInt (text);
			if (value <= 0)
			{
				showError ("Invalid " + label, label + " must be a positive integer (>= 1).");
				return null;
			}
			return value;
		}
		catch (NumberFormatException e)
		{
			showError ("Invalid " + label, label + " must be a valid integer.");
			return null;
		}
	}

	@Override
	public boolean onLeave ()
	{
		Integer headerRow = parseOptionalRowNumber (headerRowText.getText ().trim ());
		Integer dataStartRow = parseRequiredRowNumber (dataStartRowText.getText ().trim (), "Data start row");

		if (dataStartRow == null)
		{
			return false;
		}

		if (headerRow != null && headerRow > 0)
		{
			if (dataStartRow <= headerRow)
			{
				showError ("Invalid row configuration",
				           "Data start row must be greater than header row.\n\n" + "Current values:\n" +
				                                        "Header row: " + headerRow + "\n" + "Data start row: " +
				                                        dataStartRow);
				return false;
			}
		}

		String tableName = (tableNameText != null)? tableNameText.getText ().trim () : "";
		if (tableName.isEmpty ())
		{
			showError ("Invalid table name", "You must specify a table name before continuing.");
			return false;
		}

		List <Boolean> fillColumnList = new ArrayList <> ();
		if (fillEmptyPerColumn != null)
		{
			for (boolean b : fillEmptyPerColumn)
			{
				fillColumnList.add (b);
			}
		}

		// Guardar en configuración común
		config.setHeaderRow (headerRow);
		config.setDataStartRow (dataStartRow);
		config.setAutoIncrement (autoIncrementCheckbox.getSelection ());
		config.setSheetNames (new ArrayList <> (sheetNames));
		config.setTableName (tableName);
		config.setFillEmptyColumns (fillColumnList);

		return true;
	}

	@Override
	public boolean canGoNext ()
	{
		// Validación real se hace en onLeave()
		return true;
	}

	@Override
	public boolean canGoBack ()
	{
		return true;
	}

	private void showWarning (String title, String message)
	{
		MessageBox mb = new MessageBox (control.getShell (), SWT.ICON_WARNING | SWT.OK);
		mb.setText (title);
		mb.setMessage (message);
		mb.open ();
	}

	private void showError (String title, String message)
	{
		MessageBox mb = new MessageBox (control.getShell (), SWT.ICON_ERROR | SWT.OK);
		mb.setText (title);
		mb.setMessage (message);
		mb.open ();
	}

	// === Public getters (si los sigues usando en otros steps) ===

	public Integer getHeaderRow ()
	{
		return config.getHeaderRow ();
	}

	public Integer getDataStartRow ()
	{
		return config.getDataStartRow ();
	}

	public boolean isAutoIncrementEnabled ()
	{
		return autoIncrementCheckbox.getSelection ();
	}

	public File getSelectedFile ()
	{
		return selectedFile;
	}

	public List <String> getSheetNames ()
	{
		return new ArrayList <> (sheetNames);
	}

	// === Helpers ===

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

	private void toggleFillEmptyForColumn (int columnIndex)
	{
		if (fillEmptyPerColumn == null || columnIndex < 0 || columnIndex >= fillEmptyPerColumn.length)
		{
			return;
		}

		fillEmptyPerColumn[columnIndex] = !fillEmptyPerColumn[columnIndex];

		// Actualizamos texto de la "prefila"
		if (previewTable != null && !previewTable.isDisposed () && previewTable.getItemCount () > 0)
		{
			TableItem prefItem = previewTable.getItem (0);
			prefItem.setText (columnIndex + 1, fillEmptyPerColumn[columnIndex]? "☑" : "☐");
		}
	}

}
