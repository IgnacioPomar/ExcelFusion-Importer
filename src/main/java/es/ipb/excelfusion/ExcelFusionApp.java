
package es.ipb.excelfusion;

import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import es.ipb.excelfusion.ui.wizard.Step1FileSelectionPage;
import es.ipb.excelfusion.ui.wizard.Step2PreviewPage;
import es.ipb.excelfusion.ui.wizard.Step3StructureValidationPage;
import es.ipb.excelfusion.ui.wizard.Step4TypeInferencePage;
import es.ipb.excelfusion.ui.wizard.Step5DatabaseConfigPage;


public class ExcelFusionApp
{

	public static void main (String[] args)
	{

		// Create SWT display
		Display display = new Display ();

		// Main application window (will host the wizard)
		Shell shell = new Shell (display);
		shell.setText ("ExcelFusion Importer");
		shell.setSize (900, 700);

		// Initialize WizardController here
		es.ipb.excelfusion.ui.wizard.WizardController wizard = new es.ipb.excelfusion.ui.wizard.WizardController (
		        shell);
		Step1FileSelectionPage step1 = new Step1FileSelectionPage ();
		Step2PreviewPage step2 = new Step2PreviewPage (step1);
		Step3StructureValidationPage step3 = new Step3StructureValidationPage (step1, step2);
		Step4TypeInferencePage step4 = new Step4TypeInferencePage (step1, step2, step3);
		Step5DatabaseConfigPage step5 = new Step5DatabaseConfigPage ();

		wizard.addPage (step1);
		wizard.addPage (step2);
		wizard.addPage (step3);
		wizard.addPage (step4);
		wizard.addPage (step5);

		shell.open ();
		wizard.start ();

		// SWT event loop
		while (!shell.isDisposed ())
		{
			if (!display.readAndDispatch ())
			{
				display.sleep ();
			}
		}

		// Cleanup
		display.dispose ();
	}
}
