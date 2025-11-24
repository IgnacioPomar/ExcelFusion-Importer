
package es.ipb.excelfusion;

import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import es.ipb.excelfusion.config.ImportConfiguration;
import es.ipb.excelfusion.ui.wizard.Step1FileSelectionPage;
import es.ipb.excelfusion.ui.wizard.Step2PreviewPage;
import es.ipb.excelfusion.ui.wizard.Step3StructureValidationPage;
import es.ipb.excelfusion.ui.wizard.Step4TypeInferencePage;
import es.ipb.excelfusion.ui.wizard.Step5DatabaseConfigPage;
import es.ipb.excelfusion.ui.wizard.Step6ImportExecutionPage;


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

		ImportConfiguration config = new ImportConfiguration ();

		// Initialize WizardController here
		es.ipb.excelfusion.ui.wizard.WizardController wizard = new es.ipb.excelfusion.ui.wizard.WizardController (
		        shell);
		Step1FileSelectionPage step1 = new Step1FileSelectionPage (config, wizard);
		Step2PreviewPage step2 = new Step2PreviewPage (config);
		Step3StructureValidationPage step3 = new Step3StructureValidationPage (config);
		Step4TypeInferencePage step4 = new Step4TypeInferencePage (config);
		Step5DatabaseConfigPage step5 = new Step5DatabaseConfigPage (config);
		Step6ImportExecutionPage step6 = new Step6ImportExecutionPage (config);

		wizard.addPage (step1);
		wizard.addPage (step2);
		wizard.addPage (step3);
		wizard.addPage (step4);
		wizard.addPage (step5);
		wizard.addPage (step6);

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
