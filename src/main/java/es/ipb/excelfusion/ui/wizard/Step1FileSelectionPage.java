
package es.ipb.excelfusion.ui.wizard;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Locale;
import java.util.Set;

import org.eclipse.swt.SWT;
import org.eclipse.swt.events.ModifyEvent;
import org.eclipse.swt.events.ModifyListener;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.graphics.Color;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.layout.RowLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.DirectoryDialog;
import org.eclipse.swt.widgets.Group;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Table;
import org.eclipse.swt.widgets.TableColumn;
import org.eclipse.swt.widgets.TableItem;
import org.eclipse.swt.widgets.Text;

import es.ipb.excelfusion.config.ImportConfiguration;


/**
 * Step 1 of the wizard:
 * - Select data directory
 * - Choose mode (single file / auto selection)
 * - Filter files by pattern
 * - Mark already imported files (from traspasados_a_BBDD.txt)
 */
public class Step1FileSelectionPage implements WizardPage
{

	private Composite					control;

	private Text						directoryText;
	private Button						browseButton;

	private Button						singleModeRadio;
	private Button						autoModeRadio;

	private Text						filterText;
	private Table						fileTable;

	private File						currentDirectory;
	private final java.util.List <File>	excelFiles		  = new ArrayList <> ();
	private final Set <String>			importedFileNames = new HashSet <> ();

	private Color						grayColor;

	private ImportConfiguration			config;
	private final WizardController		wizardController;

	public Step1FileSelectionPage (ImportConfiguration config, WizardController wizardController)
	{
		this.config = config;
		this.wizardController = wizardController;
	}

	@Override
	public String getTitle ()
	{
		return "File selection";
	}

	@Override
	public String getDescription ()
	{
		return "Choose the directory, select Excel files to import, and optionally filter them by pattern.";
	}

	@Override
	public void createControl (Composite parent)
	{
		control = new Composite (parent, SWT.NONE);
		control.setLayout (new GridLayout (1, false));

		grayColor = parent.getDisplay ().getSystemColor (SWT.COLOR_DARK_GRAY);

		createDirectorySection (control);
		createModeSection (control);
		createFilterSection (control);
		createFileTable (control);
	}

	private void createDirectorySection (Composite parent)
	{
		Group group = new Group (parent, SWT.NONE);
		group.setText ("Data directory");
		group.setLayoutData (new GridData (SWT.FILL, SWT.TOP, true, false));
		group.setLayout (new GridLayout (3, false));

		Label dirLabel = new Label (group, SWT.NONE);
		dirLabel.setText ("Directory:");

		directoryText = new Text (group, SWT.BORDER | SWT.READ_ONLY);
		directoryText.setLayoutData (new GridData (SWT.FILL, SWT.CENTER, true, false));

		browseButton = new Button (group, SWT.PUSH);
		browseButton.setText ("Browse...");
		browseButton.addSelectionListener (new SelectionAdapter ()
		{
			@Override
			public void widgetSelected (SelectionEvent e)
			{
				onBrowseDirectory ();
			}
		});
	}

	private void createModeSection (Composite parent)
	{
		Group group = new Group (parent, SWT.NONE);
		group.setText ("Mode");
		group.setLayoutData (new GridData (SWT.FILL, SWT.TOP, true, false));
		group.setLayout (new RowLayout (SWT.HORIZONTAL));

		singleModeRadio = new Button (group, SWT.RADIO);
		singleModeRadio.setText ("Single file mode");
		singleModeRadio.setSelection (true); // default

		autoModeRadio = new Button (group, SWT.RADIO);
		autoModeRadio.setText ("Auto-selection mode");
	}

	private void createFilterSection (Composite parent)
	{
		Group group = new Group (parent, SWT.NONE);
		group.setText ("Filter (auto-selection)");
		group.setLayoutData (new GridData (SWT.FILL, SWT.TOP, true, false));
		group.setLayout (new GridLayout (2, false));

		Label lbl = new Label (group, SWT.NONE);
		lbl.setText ("Pattern (space separated words):");

		filterText = new Text (group, SWT.BORDER);
		filterText.setLayoutData (new GridData (SWT.FILL, SWT.CENTER, true, false));

		filterText.addModifyListener (new ModifyListener ()
		{
			@Override
			public void modifyText (ModifyEvent e)
			{
				applyFilter ();
			}
		});
	}

	private void createFileTable (Composite parent)
	{
		Group group = new Group (parent, SWT.NONE);
		group.setText ("Excel files");
		group.setLayoutData (new GridData (SWT.FILL, SWT.FILL, true, true));
		group.setLayout (new GridLayout (1, false));

		fileTable = new Table (group, SWT.BORDER | SWT.CHECK | SWT.FULL_SELECTION | SWT.V_SCROLL | SWT.H_SCROLL);
		fileTable.setHeaderVisible (true);
		fileTable.setLinesVisible (true);

		GridData gd = new GridData (SWT.FILL, SWT.FILL, true, true);
		gd.heightHint = 300;
		fileTable.setLayoutData (gd);

		TableColumn colName = new TableColumn (fileTable, SWT.LEFT);
		colName.setText ("File");
		colName.setWidth (350);

		TableColumn colStatus = new TableColumn (fileTable, SWT.LEFT);
		colStatus.setText ("Status");
		colStatus.setWidth (120);

		fileTable.addSelectionListener (new SelectionAdapter ()
		{
			@Override
			public void widgetSelected (SelectionEvent e)
			{
				if (e.detail == SWT.CHECK)
				{
					onFileChecked ((TableItem) e.item);
				}
			}
		});
	}

	private void onBrowseDirectory ()
	{
		DirectoryDialog dialog = new DirectoryDialog (control.getShell ());
		dialog.setMessage ("Select the directory containing Excel files (.xls, .xlsx)");
		String dir = dialog.open ();
		if (dir != null)
		{
			currentDirectory = new File (dir);
			directoryText.setText (currentDirectory.getAbsolutePath ());
			loadImportedFileList ();
			loadExcelFiles ();
			applyFilter ();
		}
	}

	private void loadImportedFileList ()
	{
		importedFileNames.clear ();
		if (currentDirectory == null)
		{
			return;
		}
		File importedFile = new File (currentDirectory, "traspasados_a_BBDD.txt");
		if (!importedFile.exists ())
		{
			return;
		}

		try (BufferedReader br = new BufferedReader (new FileReader (importedFile)))
		{
			String line;
			while ((line = br.readLine ()) != null)
			{
				String trimmed = line.trim ();
				if (!trimmed.isEmpty ())
				{
					importedFileNames.add (trimmed);
				}
			}
		}
		catch (IOException e)
		{
			// For now, just print stack trace; later replace with proper logging
			e.printStackTrace ();
		}
	}

	private void loadExcelFiles ()
	{
		excelFiles.clear ();
		fileTable.removeAll ();

		if (currentDirectory == null || !currentDirectory.isDirectory ())
		{
			return;
		}

		File[] files = currentDirectory.listFiles ( (dir, name) -> {
			String lower = name.toLowerCase (Locale.ROOT);
			return lower.endsWith (".xls") || lower.endsWith (".xlsx");
		});

		if (files == null)
		{
			return;
		}

		excelFiles.addAll (Arrays.asList (files));
		for (File f : excelFiles)
		{
			TableItem item = new TableItem (fileTable, SWT.NONE);
			item.setText (0, f.getName ());

			boolean imported = importedFileNames.contains (f.getName ());
			if (imported)
			{
				item.setText (1, "Already imported");
				item.setForeground (grayColor);
				item.setData ("imported", Boolean.TRUE);
				item.setChecked (false);
			}
			else
			{
				item.setText (1, "");
				item.setData ("imported", Boolean.FALSE);
			}
		}
	}

	private void applyFilter ()
	{
		if (fileTable.isDisposed ())
		{
			return;
		}

		String pattern = filterText.getText ().trim ().toLowerCase (Locale.ROOT);
		java.util.List <String> words = new ArrayList <> ();
		if (!pattern.isEmpty ())
		{
			for (String w : pattern.split ("\\s+"))
			{
				if (!w.isEmpty ())
				{
					words.add (w);
				}
			}
		}

		TableItem[] items = fileTable.getItems ();
		for (TableItem item : items)
		{
			String fileName = item.getText (0).toLowerCase (Locale.ROOT);
			boolean imported = Boolean.TRUE.equals (item.getData ("imported"));

			boolean matches = true;
			for (String w : words)
			{
				if (!fileName.contains (w))
				{
					matches = false;
					break;
				}
			}

			if (imported)
			{
				// imported: always gray and non-selectable
				item.setForeground (grayColor);
				item.setChecked (false);
			}
			else if (!words.isEmpty () && !matches)
			{
				// not matching filter: gray
				item.setForeground (grayColor);
			}
			else
			{
				// matches filter (or empty pattern)
				item.setForeground (null); // default color
			}
		}
	}

	private void onFileChecked (TableItem changedItem)
	{
		boolean imported = Boolean.TRUE.equals (changedItem.getData ("imported"));
		if (imported)
		{
			// Never allow checking imported files
			changedItem.setChecked (false);
			return;
		}

		if (singleModeRadio.getSelection ())
		{
			// In single file mode, uncheck all others
			for (TableItem item : fileTable.getItems ())
			{
				if (item != changedItem)
				{
					item.setChecked (false);
				}
			}
		}

		// refrescar botones del wizard
		if (wizardController != null)
		{
			wizardController.refreshButtons ();
		}

	}

	@Override
	public Composite getControl ()
	{
		return control;
	}

	@Override
	public void onEnter ()
	{
		// If no directory selected yet, user must browse.
		// Could auto-restore last directory from config later.
	}

	@Override
	public boolean onLeave ()
	{
		// Ensure at least one file is selected before going next
		if (!hasSelectedFiles ())
		{
			MessageBox mb = new MessageBox (control.getShell (), SWT.ICON_WARNING | SWT.OK);
			mb.setText ("No files selected");
			mb.setMessage ("Please select at least one Excel file before continuing.");
			mb.open ();
			return false;
		}

		// Make sure the fikes are accessible
		for (File f : getSelectedFiles ())
		{
			if (!f.exists () || !f.canRead ())
			{
				MessageBox mb = new MessageBox (control.getShell (), SWT.ICON_ERROR | SWT.OK);
				mb.setText ("File access error");
				mb.setMessage ("The file \"" + f.getName () +
				               "\" cannot be accessed. Please check that it exists and is readable.");
				mb.open ();
				return false;
			}
		}

		config.setDataDirectory (currentDirectory);
		config.setSelectedFiles (getSelectedFiles ());

		return true;
	}

	@Override
	public boolean canGoNext ()
	{
		// Dynamically: only allow "Next" if some file is selected
		return hasSelectedFiles ();
	}

	@Override
	public boolean canGoBack ()
	{
		// This is the first page, so usually false
		return false;
	}

	private boolean hasSelectedFiles ()
	{
		if (fileTable == null || fileTable.isDisposed ())
		{
			return false;
		}
		for (TableItem item : fileTable.getItems ())
		{
			boolean imported = Boolean.TRUE.equals (item.getData ("imported"));
			if (!imported && item.getChecked ())
			{
				return true;
			}
		}
		return false;
	}

	/**
	 * Returns the selected files.
	 * In single file mode: max 1 file.
	 * In auto-selection mode: possibly multiple files.
	 */
	public java.util.List <File> getSelectedFiles ()
	{
		java.util.List <File> selected = new ArrayList <> ();
		if (currentDirectory == null)
		{
			return selected;
		}
		TableItem[] items = fileTable.getItems ();
		for (int i = 0; i < items.length; i++)
		{
			TableItem item = items[i];
			boolean imported = Boolean.TRUE.equals (item.getData ("imported"));
			if (!imported && item.getChecked ())
			{
				selected.add (excelFiles.get (i));
			}
		}
		return selected;
	}

	public File getCurrentDirectory ()
	{
		return currentDirectory;
	}
}
