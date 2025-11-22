
package es.ipb.excelfusion.ui.wizard;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.eclipse.swt.SWT;
import org.eclipse.swt.graphics.Color;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Group;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Table;
import org.eclipse.swt.widgets.TableColumn;
import org.eclipse.swt.widgets.TableItem;

import es.ipb.excelfusion.config.ImportConfiguration;


/**
 * Step 3 of the wizard:
 * - If a header row is defined:
 * - Validate that all selected files and sheets share the same header structure.
 * - Show per file@sheet status (OK / Mismatch / Error).
 * - Optionally import only matching sheets.
 *
 * - If no header row is defined:
 * - Show a warning that structure cannot be validated.
 * - Allow user to continue using generic column names (A, B, C, ...).
 */
public class Step3StructureValidationPage implements WizardPage
{

	private Composite							   control;

	private Label								   infoLabel;
	private Table								   resultTable;
	private Button								   importOnlyMatchingCheckbox;

	private boolean								   headerDefined		 = false;
	private boolean								   validationDone		 = false;
	private boolean								   hasMismatchesOrErrors = false;

	private java.util.List <SheetValidationResult> results				 = new ArrayList <> ();

	private Color								   okColor;
	private Color								   errorColor;

	private ImportConfiguration					   config;

	public Step3StructureValidationPage (ImportConfiguration config)
	{
		this.config = config;
	}

	@Override
	public String getTitle ()
	{
		return "Structure validation";
	}

	@Override
	public String getDescription ()
	{
		return "Check that all selected files and sheets share the same header structure.";
	}

	@Override
	public void createControl (Composite parent)
	{
		control = new Composite (parent, SWT.NONE);
		control.setLayout (new GridLayout (1, false));

		okColor = parent.getDisplay ().getSystemColor (SWT.COLOR_DARK_GREEN);
		errorColor = parent.getDisplay ().getSystemColor (SWT.COLOR_DARK_RED);

		createInfoSection (control);
		createResultTable (control);
		createOptionsSection (control);
	}

	private void createInfoSection (Composite parent)
	{
		infoLabel = new Label (parent, SWT.WRAP);
		infoLabel.setLayoutData (new GridData (SWT.FILL, SWT.TOP, true, false));
		infoLabel.setText ("Validation has not been executed yet.");
	}

	private void createResultTable (Composite parent)
	{
		Group group = new Group (parent, SWT.NONE);
		group.setText ("Header validation results");
		group.setLayoutData (new GridData (SWT.FILL, SWT.FILL, true, true));
		group.setLayout (new GridLayout (1, false));

		resultTable = new Table (group, SWT.BORDER | SWT.FULL_SELECTION | SWT.V_SCROLL | SWT.H_SCROLL);
		resultTable.setHeaderVisible (true);
		resultTable.setLinesVisible (true);

		GridData gd = new GridData (SWT.FILL, SWT.FILL, true, true);
		gd.heightHint = 300;
		resultTable.setLayoutData (gd);

		TableColumn colFile = new TableColumn (resultTable, SWT.LEFT);
		colFile.setText ("File");
		colFile.setWidth (250);

		TableColumn colSheet = new TableColumn (resultTable, SWT.LEFT);
		colSheet.setText ("Sheet");
		colSheet.setWidth (150);

		TableColumn colStatus = new TableColumn (resultTable, SWT.LEFT);
		colStatus.setText ("Status");
		colStatus.setWidth (120);

		TableColumn colMessage = new TableColumn (resultTable, SWT.LEFT);
		colMessage.setText ("Details");
		colMessage.setWidth (300);
	}

	private void createOptionsSection (Composite parent)
	{
		Composite bottom = new Composite (parent, SWT.NONE);
		bottom.setLayoutData (new GridData (SWT.FILL, SWT.TOP, true, false));
		bottom.setLayout (new GridLayout (1, false));

		importOnlyMatchingCheckbox = new Button (bottom, SWT.CHECK);
		importOnlyMatchingCheckbox.setText ("Import only matching files/sheets");
		importOnlyMatchingCheckbox.setSelection (true);
		importOnlyMatchingCheckbox.setVisible (false); // only shown if mismatches exist
	}

	@Override
	public Composite getControl ()
	{
		return control;
	}

	@Override
	public void onEnter ()
	{
		// Run validation only once per visit (unless you want to re-run always)
		if (!validationDone)
		{
			runValidation ();
		}
	}

	private void runValidation ()
	{
		results.clear ();
		resultTable.removeAll ();
		hasMismatchesOrErrors = false;

		Integer headerRow = config.getHeaderRow (); // 1-based, null if none
		if (headerRow == null || headerRow <= 0)
		{
			// No header defined
			headerDefined = false;
			infoLabel.setText ("No header row has been defined.\n\n" +
			                   "It is not possible to validate the structure of the sheets and files.\n" +
			                   "The import will use generic column names (A, B, C, ...).");
			importOnlyMatchingCheckbox.setVisible (false);
			validationDone = true;
			return;
		}

		headerDefined = true;

		java.util.List <File> selectedFiles = config.getSelectedFiles ();
		if (selectedFiles.isEmpty ())
		{
			infoLabel.setText ("No selected files found. Please go back and select at least one file.");
			importOnlyMatchingCheckbox.setVisible (false);
			validationDone = true;
			return;
		}

		// 0-based index for POI (headerRow is 1-based)
		int headerRowIndex = headerRow - 1;
		DataFormatter formatter = new DataFormatter (Locale.getDefault ());

		java.util.List <String> referenceHeader = null;
		File referenceFile = null;
		String referenceSheetName = null;

		// Iterate over all files and all sheets
		for (File file : selectedFiles)
		{
			try (FileInputStream fis = new FileInputStream (file); Workbook workbook = WorkbookFactory.create (fis))
			{

				int numberOfSheets = workbook.getNumberOfSheets ();

				for (int i = 0; i < numberOfSheets; i++)
				{
					Sheet sheet = workbook.getSheetAt (i);
					String sheetName = sheet.getSheetName ();

					SheetValidationResult result = new SheetValidationResult (file, sheetName);

					try
					{
						Row row = sheet.getRow (headerRowIndex);
						java.util.List <String> headerValues = new ArrayList <> ();

						if (row != null)
						{
							short lastCellNum = row.getLastCellNum ();
							if (lastCellNum < 0)
							{
								// row exists but has no cells
								lastCellNum = 0;
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
						}

						if (referenceHeader == null)
						{
							// First sheet: define reference
							referenceHeader = headerValues;
							referenceFile = file;
							referenceSheetName = sheetName;

							result.setMatches (true);
							result.setStatusMessage ("Reference header");
						}
						else
						{
							boolean matches = headersEqual (referenceHeader, headerValues);
							result.setMatches (matches);
							if (matches)
							{
								result.setStatusMessage ("OK");
							}
							else
							{
								result.setStatusMessage ("Header does not match reference");
								hasMismatchesOrErrors = true;
							}
						}
					}
					catch (Exception exSheet)
					{
						result.setMatches (false);
						result.setStatusMessage ("Error reading header: " + exSheet.getMessage ());
						hasMismatchesOrErrors = true;
					}

					results.add (result);
				}

			}
			catch (IOException ioEx)
			{
				SheetValidationResult errorResult = new SheetValidationResult (file, "<all sheets>");
				errorResult.setMatches (false);
				errorResult.setStatusMessage ("Error opening file: " + ioEx.getMessage ());
				results.add (errorResult);
				hasMismatchesOrErrors = true;
			}
			catch (Exception ex)
			{
				SheetValidationResult errorResult = new SheetValidationResult (file, "<all sheets>");
				errorResult.setMatches (false);
				errorResult.setStatusMessage ("Unexpected error: " + ex.getMessage ());
				results.add (errorResult);
				hasMismatchesOrErrors = true;
			}
		}

		// Update UI table
		fillResultTable ();

		// Update info label
		if (referenceHeader == null)
		{
			infoLabel.setText ("Could not determine a reference header (all header rows seem empty or invalid).\n" +
			                   "Please check your header row configuration.");
		}
		else
		{
			StringBuilder sb = new StringBuilder ();
			sb.append ("Header structure validation completed.\n");
			sb.append ("Reference: ").append (referenceFile != null? referenceFile.getName () : "?").append (" @ ")
			        .append (referenceSheetName != null? referenceSheetName : "?").append ("\n\n");

			if (hasMismatchesOrErrors)
			{
				sb.append ("Some sheets do not match the reference header or produced errors.\n");
				sb.append ("You may choose to import only matching files/sheets.");
			}
			else
			{
				sb.append ("All sheets in all selected files match the reference header.\n");
				sb.append ("Everything is ready for import.");
			}

			infoLabel.setText (sb.toString ());
		}

		importOnlyMatchingCheckbox.setVisible (hasMismatchesOrErrors);
		importOnlyMatchingCheckbox.setSelection (hasMismatchesOrErrors);

		validationDone = true;
	}

	private boolean headersEqual (java.util.List <String> reference, java.util.List <String> candidate)
	{
		if (reference == null && candidate == null)
		{
			return true;
		}
		if (reference == null || candidate == null)
		{
			return false;
		}
		if (reference.size () != candidate.size ())
		{
			return false;
		}
		for (int i = 0; i < reference.size (); i++)
		{
			String a = safeString (reference.get (i)).trim ();
			String b = safeString (candidate.get (i)).trim ();
			if (!a.equals (b))
			{
				return false;
			}
		}
		return true;
	}

	private String safeString (String s)
	{
		return (s == null)? "" : s;
	}

	private void fillResultTable ()
	{
		resultTable.removeAll ();

		for (SheetValidationResult r : results)
		{
			TableItem item = new TableItem (resultTable, SWT.NONE);
			item.setText (0, r.getFile ().getName ());
			item.setText (1, r.getSheetName ());

			if (r.isMatches ())
			{
				item.setText (2, "OK");
				item.setText (3, r.getStatusMessage () != null? r.getStatusMessage () : "");
				item.setForeground (okColor);
			}
			else
			{
				item.setText (2, "Mismatch/Error");
				item.setText (3, r.getStatusMessage () != null? r.getStatusMessage () : "");
				item.setForeground (errorColor);
			}
		}
	}

	@Override
	public boolean onLeave ()
	{
		// No extra validation here; we just keep the results.
		return true;
	}

	@Override
	public boolean canGoNext ()
	{
		// You could decide to block "Next" if headerDefined and referenceHeader == null
		return true;
	}

	@Override
	public boolean canGoBack ()
	{
		return true;
	}

	/**
	 * Returns the list of file@sheet combinations that should be imported,
	 * depending on the "Import only matching" checkbox.
	 */
	public java.util.List <SheetValidationResult> getSheetsToImport ()
	{
		java.util.List <SheetValidationResult> selected = new ArrayList <> ();

		boolean onlyMatching = importOnlyMatchingCheckbox.getVisible () && importOnlyMatchingCheckbox.getSelection ();

		for (SheetValidationResult r : results)
		{
			if (!headerDefined)
			{
				// If no header defined, we don't actually filter per sheet here.
				// Later steps will decide how to handle generic columns.
				// For now, include everything that isn't a fatal file-level error.
				selected.add (r);
			}
			else
			{
				if (onlyMatching)
				{
					if (r.isMatches ())
					{
						selected.add (r);
					}
				}
				else
				{
					// Include everything, even mismatches/errors
					selected.add (r);
				}
			}
		}

		return selected;
	}

	public boolean isHeaderDefined ()
	{
		return headerDefined;
	}

	public boolean hasMismatchesOrErrors ()
	{
		return hasMismatchesOrErrors;
	}

	/**
	 * Simple holder for per-sheet validation result.
	 */
	public static class SheetValidationResult
	{

		private final File	 file;
		private final String sheetName;
		private boolean		 matches;
		private String		 statusMessage;

		public SheetValidationResult (File file, String sheetName)
		{
			this.file = file;
			this.sheetName = sheetName;
		}

		public File getFile ()
		{
			return file;
		}

		public String getSheetName ()
		{
			return sheetName;
		}

		public boolean isMatches ()
		{
			return matches;
		}

		public void setMatches (boolean matches)
		{
			this.matches = matches;
		}

		public String getStatusMessage ()
		{
			return statusMessage;
		}

		public void setStatusMessage (String statusMessage)
		{
			this.statusMessage = statusMessage;
		}
	}
}
