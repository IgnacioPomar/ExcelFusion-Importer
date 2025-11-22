
package es.ipb.excelfusion.ui.wizard;

import java.util.ArrayList;
import java.util.List;

import org.eclipse.swt.SWT;
import org.eclipse.swt.custom.StackLayout;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.layout.RowLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Shell;


/**
 * Controls the multi-step wizard:
 * - Hosts all WizardPages inside a StackLayout
 * - Handles Next / Back navigation
 * - Updates button states based on the current page
 */
public class WizardController
{

	private final Shell				shell;
	private Composite				pageContainer;
	private Composite				buttonBar;

	private Button					backButton;
	private Button					nextButton;

	private final List <WizardPage>	pages			 = new ArrayList <> ();
	private int						currentPageIndex = -1;

	private final StackLayout		stackLayout		 = new StackLayout ();

	public WizardController (Shell shell)
	{
		this.shell = shell;
		createLayout ();
	}

	/**
	 * Creates the base layout:
	 * - pageContainer (center, with StackLayout)
	 * - buttonBar (bottom, with Back / Next)
	 */
	private void createLayout ()
	{
		GridLayout rootLayout = new GridLayout (1, false);
		rootLayout.marginWidth = 10;
		rootLayout.marginHeight = 10;
		rootLayout.verticalSpacing = 10;
		shell.setLayout (rootLayout);

		// Container for pages
		pageContainer = new Composite (shell, SWT.NONE);
		pageContainer.setLayoutData (new GridData (SWT.FILL, SWT.FILL, true, true));
		pageContainer.setLayout (stackLayout);

		// Simple placeholder while there are no pages
		Label placeholder = new Label (pageContainer, SWT.CENTER);
		placeholder.setText ("No pages registered yet.");
		stackLayout.topControl = placeholder;

		// Button bar
		buttonBar = new Composite (shell, SWT.NONE);
		GridData buttonBarData = new GridData (SWT.FILL, SWT.CENTER, true, false);
		buttonBar.setLayoutData (buttonBarData);

		RowLayout rowLayout = new RowLayout (SWT.HORIZONTAL);
		rowLayout.justify = true;
		rowLayout.fill = false;
		rowLayout.marginWidth = 0;
		rowLayout.marginHeight = 0;
		rowLayout.spacing = 10;
		buttonBar.setLayout (rowLayout);

		backButton = new Button (buttonBar, SWT.PUSH);
		backButton.setText ("< Back");
		backButton.addSelectionListener (new SelectionAdapter ()
		{
			@Override
			public void widgetSelected (SelectionEvent e)
			{
				showPreviousPage ();
			}
		});

		nextButton = new Button (buttonBar, SWT.PUSH);
		nextButton.setText ("Next >");
		nextButton.addSelectionListener (new SelectionAdapter ()
		{
			@Override
			public void widgetSelected (SelectionEvent e)
			{
				handleNextPressed ();
			}
		});

		updateButtonState ();
	}

	/**
	 * Register a new wizard page.
	 */
	public void addPage (WizardPage page)
	{
		pages.add (page);
	}

	/**
	 * Show the first page (index 0). Call this after registering all pages.
	 */
	public void start ()
	{
		if (pages.isEmpty ())
		{
			return;
		}
		showPage (0);
	}

	private void handleNextPressed ()
	{
		if (currentPageIndex < 0 || currentPageIndex >= pages.size ())
		{
			return;
		}

		WizardPage current = pages.get (currentPageIndex);

		// Allow the page to validate or block navigation
		if (!current.onLeave ())
		{
			return; // stay on this page
		}

		boolean isLastPage = currentPageIndex == pages.size () - 1;
		if (isLastPage)
		{
			// Finish the wizard
			shell.close ();
		}
		else
		{
			showPage (currentPageIndex + 1);
		}
	}

	private void showPreviousPage ()
	{
		if (currentPageIndex <= 0)
		{
			return;
		}

		WizardPage current = pages.get (currentPageIndex);
		if (!current.onLeave ())
		{
			return; // optional: block going backwards if page says so
		}

		showPage (currentPageIndex - 1);
	}

	/**
	 * Show page by index, creating its control if needed.
	 */
	private void showPage (int index)
	{
		if (index < 0 || index >= pages.size ())
		{
			return;
		}

		WizardPage page = pages.get (index);

		if (page.getControl () == null || page.getControl ().isDisposed ())
		{
			page.createControl (pageContainer);
		}

		currentPageIndex = index;

		// Update shell title with page title (optional)
		String baseTitle = "ExcelFusion Importer";
		if (page.getTitle () != null && !page.getTitle ().isBlank ())
		{
			shell.setText (baseTitle + " - " + page.getTitle ());
		}
		else
		{
			shell.setText (baseTitle);
		}

		stackLayout.topControl = page.getControl ();
		pageContainer.layout ();

		page.onEnter ();
		updateButtonState ();
	}

	/**
	 * Enable/disable Back/Next depending on current page and page rules.
	 */
	private void updateButtonState ()
	{
		if (pages.isEmpty () || currentPageIndex < 0)
		{
			backButton.setEnabled (false);
			nextButton.setEnabled (false);
			return;
		}

		WizardPage current = pages.get (currentPageIndex);

		boolean isFirst = currentPageIndex == 0;
		boolean isLast = currentPageIndex == pages.size () - 1;

		backButton.setEnabled (!isFirst && current.canGoBack ());
		nextButton.setEnabled (current.canGoNext ());

		if (isLast)
		{
			nextButton.setText ("Finish");
		}
		else
		{
			nextButton.setText ("Next >");
		}

		buttonBar.layout ();
	}
}



/**
 * Simple contract for wizard pages.
 * You can move this to its own file (WizardPage.java) if you prefer.
 */
interface WizardPage
{

	/**
	 * Human-readable title for the page (used in shell title).
	 */
	String getTitle ();

	/**
	 * (Optional) Short description or subtitle.
	 */
	default String getDescription ()
	{
		return "";
	}

	/**
	 * Called once to create the page contents.
	 */
	void createControl (Composite parent);

	/**
	 * Returns the root control for this page.
	 */
	Composite getControl ();

	/**
	 * Called when the page becomes visible.
	 */
	default void onEnter ()
	{
		// no-op by default
	}

	/**
	 * Called before leaving the page.
	 * Return false to prevent navigation.
	 */
	default boolean onLeave ()
	{
		return true;
	}

	/**
	 * Whether the wizard can go to the next page from here.
	 */
	default boolean canGoNext ()
	{
		return true;
	}

	/**
	 * Whether the wizard can go back from this page.
	 */
	default boolean canGoBack ()
	{
		return true;
	}
}
