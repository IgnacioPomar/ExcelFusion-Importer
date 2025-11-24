
package es.ipb.excelfusion.ui.wizard;

import java.io.File;

import org.eclipse.swt.SWT;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.layout.RowLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Group;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Text;

import es.ipb.excelfusion.config.ImportConfiguration;
import es.ipb.excelfusion.service.ImportExecutor;
import es.ipb.excelfusion.service.ImportProgressListener;


/**
 * Wizard step 6:
 * - Pure SWT UI for import execution.
 * - Delegates all business logic to ImportExecutor (no DB/Excel logic here).
 */
public class Step6ImportExecutionPage implements WizardPage
{

	private final ImportConfiguration config;

	private Composite				  control;
	private Text					  logText;
	private Label					  progressLabel;
	private Button					  startButton;

	private volatile boolean		  importRunning	 = false;
	private volatile boolean		  importFinished = false;

	public Step6ImportExecutionPage (ImportConfiguration config)
	{
		this.config = config;
	}

	@Override
	public String getTitle ()
	{
		return "Import execution";
	}

	@Override
	public String getDescription ()
	{
		return "Run the import process and watch progress and log output.";
	}

	@Override
	public void createControl (Composite parent)
	{
		control = new Composite (parent, SWT.NONE);
		control.setLayout (new GridLayout (1, false));

		createTopInfo (control);
		createLogArea (control);
		createBottomControls (control);
	}

	private void createTopInfo (Composite parent)
	{
		Label info = new Label (parent, SWT.WRAP);
		info.setLayoutData (new GridData (SWT.FILL, SWT.TOP, true, false));
		info.setText ("When you click 'Start import', the tool will execute the import using the current configuration.\n" +
		              "All progress and messages will be displayed below.");
	}

	private void createLogArea (Composite parent)
	{
		Group group = new Group (parent, SWT.NONE);
		group.setText ("Import log");
		group.setLayoutData (new GridData (SWT.FILL, SWT.FILL, true, true));
		group.setLayout (new GridLayout (1, false));

		logText = new Text (group, SWT.BORDER | SWT.MULTI | SWT.READ_ONLY | SWT.V_SCROLL | SWT.H_SCROLL);
		GridData gd = new GridData (SWT.FILL, SWT.FILL, true, true);
		gd.heightHint = 350;
		logText.setLayoutData (gd);

		progressLabel = new Label (group, SWT.NONE);
		progressLabel.setLayoutData (new GridData (SWT.FILL, SWT.TOP, true, false));
		progressLabel.setText ("Ready to start.");
	}

	private void createBottomControls (Composite parent)
	{
		Composite bottom = new Composite (parent, SWT.NONE);
		bottom.setLayoutData (new GridData (SWT.FILL, SWT.TOP, true, false));
		RowLayout layout = new RowLayout (SWT.HORIZONTAL);
		layout.spacing = 10;
		layout.marginTop = 5;
		layout.marginBottom = 5;
		bottom.setLayout (layout);

		startButton = new Button (bottom, SWT.PUSH);
		startButton.setText ("Start import");
		startButton.addSelectionListener (new org.eclipse.swt.events.SelectionAdapter ()
		{
			@Override
			public void widgetSelected (org.eclipse.swt.events.SelectionEvent e)
			{
				onStartImport ();
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
		// Nothing special; user has to click "Start import"
	}

	@Override
	public boolean onLeave ()
	{
		// You might want to prevent leaving while import is running
		return !importRunning;
	}

	@Override
	public boolean canGoNext ()
	{
		// Last step; wizard will close on Finish.
		// If quieres que sólo se pueda terminar tras éxito:
		// return importFinished;
		return false;
	}

	@Override
	public boolean canGoBack ()
	{
		return !importRunning;
	}

	// === UI → ImportExecutor ===

	private void onStartImport ()
	{
		if (importRunning)
		{
			return;
		}
		importRunning = true;
		importFinished = false;
		startButton.setEnabled (false);
		clearLog ();
		appendLog ("Starting import...\n");

		ImportProgressListener listener = createProgressListener ();

		Thread t = new Thread ( () -> {
			try
			{
				ImportExecutor executor = new ImportExecutor (config, listener);
				executor.execute ();
			}
			catch (Exception e)
			{
				// ImportExecutor already notified listener.onError(e)
				// Nothing else needed here.
			}
			finally
			{
				asyncExec ( () -> {
					importRunning = false;
					startButton.setEnabled (true);
				});
			}
		}, "ImportExecutorThread");

		t.setDaemon (true);
		t.start ();
	}

	private ImportProgressListener createProgressListener ()
	{
		return new ImportProgressListener ()
		{
			@Override
			public void onLog (String message)
			{
				appendLog (message);
			}

			@Override
			public void onSheetStarted (int fileIndex, int totalFiles, int sheetIndex, int totalSheets, File file,
			                            String sheetName)
			{
				String msg = "Processing " + file.getName () + "@" + sheetName + "...";
				updateProgress ("File " + fileIndex + " of " + totalFiles + " / Sheet " + sheetIndex + " of " +
				                totalSheets);
				appendLog (msg + "\n");
			}

			@Override
			public void onSheetCompleted (int fileIndex, int totalFiles, int sheetIndex, int totalSheets, File file,
			                              String sheetName)
			{
				String msg = "Completed " + file.getName () + "@" + sheetName;
				updateProgress ("File " + fileIndex + " of " + totalFiles + " / Sheet " + sheetIndex + " of " +
				                totalSheets);
				appendLog (msg + "\n");
			}

			@Override
			public void onCompleted ()
			{
				importFinished = true;
				appendLog ("Import completed successfully.\n");
				asyncExec ( () -> {
					MessageBox mb = new MessageBox (control.getShell (), SWT.ICON_INFORMATION | SWT.OK);
					mb.setText ("Import completed");
					mb.setMessage ("The import finished successfully.");
					mb.open ();
				});
			}

			@Override
			public void onError (Exception e)
			{
				appendLog ("ERROR: " + e.getMessage () + "\n");
				asyncExec ( () -> {
					MessageBox mb = new MessageBox (control.getShell (), SWT.ICON_ERROR | SWT.OK);
					mb.setText ("Import failed");
					mb.setMessage ("An error occurred during import:\n\n" + e.getMessage ());
					mb.open ();
				});
			}
		};
	}

	// === Helpers de UI ===

	private void appendLog (String text)
	{
		asyncExec ( () -> {
			if (logText != null && !logText.isDisposed ())
			{
				logText.append (text);
			}
		});
	}

	private void clearLog ()
	{
		asyncExec ( () -> {
			if (logText != null && !logText.isDisposed ())
			{
				logText.setText ("");
			}
		});
	}

	private void updateProgress (String msg)
	{
		asyncExec ( () -> {
			if (progressLabel != null && !progressLabel.isDisposed ())
			{
				progressLabel.setText (msg);
			}
		});
	}

	private void asyncExec (Runnable r)
	{
		if (control == null || control.isDisposed ())
		{
			return;
		}
		Display display = control.getDisplay ();
		if (display == null || display.isDisposed ())
		{
			return;
		}
		display.asyncExec (r);
	}
}
