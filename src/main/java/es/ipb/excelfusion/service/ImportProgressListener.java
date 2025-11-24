
package es.ipb.excelfusion.service;

import java.io.File;


/**
 * Listener for import progress notifications.
 * Implementations can update a GUI, log to console, etc.
 */
public interface ImportProgressListener
{

	/**
	 * General log message (free text).
	 */
	void onLog (String message);

	/**
	 * Called when starting to process a specific file@sheet.
	 */
	void onSheetStarted (int fileIndex, int totalFiles, int sheetIndex, int totalSheets, File file, String sheetName);

	/**
	 * Called when a specific file@sheet has been processed successfully.
	 */
	void onSheetCompleted (int fileIndex, int totalFiles, int sheetIndex, int totalSheets, File file, String sheetName);

	/**
	 * Called when the entire import has completed successfully.
	 */
	void onCompleted ();

	/**
	 * Called when a fatal error occurs and the import cannot continue.
	 */
	void onError (Exception e);
}
