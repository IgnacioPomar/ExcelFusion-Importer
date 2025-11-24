
package es.ipb.excelfusion.service;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Types;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.Locale;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import es.ipb.excelfusion.config.ImportConfiguration;
import es.ipb.excelfusion.ui.wizard.Step3StructureValidationPage.SheetValidationResult;
import es.ipb.excelfusion.ui.wizard.Step4TypeInferencePage.ColumnDefinition;
import es.ipb.excelfusion.ui.wizard.Step5DatabaseConfigPage.DbType;


/**
 * Executes the full import process based on ImportConfiguration.
 * This class contains NO UI code. It reports progress via ImportProgressListener.
 *
 * For simplicity, all data columns are created as TEXT in the DB.
 */
public class ImportExecutor
{

	private final ImportConfiguration	 config;
	private final ImportProgressListener listener;

	public ImportExecutor (ImportConfiguration config, ImportProgressListener listener)
	{
		this.config = config;
		this.listener = listener;
	}

	public void execute () throws Exception
	{
		try
		{
			doExecute ();
			notifyCompleted ();
		}
		catch (Exception e)
		{
			notifyError (e);
			throw e;
		}
	}

	private void doExecute () throws Exception
	{
		validateConfiguration ();

		java.util.List <SheetValidationResult> sheetsToImport = config.getSheetsToImport ();
		if (sheetsToImport == null || sheetsToImport.isEmpty ())
		{
			throw new IllegalStateException ("No sheets to import. Please review previous steps.");
		}

		java.util.Map <File, java.util.List <SheetValidationResult>> sheetsByFile = groupSheetsByFile (sheetsToImport);
		int totalFiles = sheetsByFile.size ();

		DbType dbType = config.getDbType ();
		String host = config.getDbHost ();
		int port = config.getDbPort ();
		String dbName = config.getDbName ();
		String user = config.getDbUser ();
		String password = config.getDbPassword ();
		String tableName = config.getTableName ();

		if (tableName == null || tableName.trim ().isEmpty ())
		{
			throw new IllegalStateException ("Table name is not defined.");
		}

		String normalizedTableName = normalizeIdentifier (tableName);
		String jdbcUrl = buildJdbcUrl (dbType, host, port, dbName);

		log ("Connecting to database: " + jdbcUrl);

		try (Connection conn = DriverManager.getConnection (jdbcUrl, user, password))
		{
			conn.setAutoCommit (false);

			if (tableExists (conn, normalizedTableName))
			{
				log ("Target table '" + normalizedTableName + "' already exists. Aborting.");
				throw new IllegalStateException ("Target table '" + normalizedTableName + "' already exists.");
			}

			log ("Creating table '" + normalizedTableName + "'...");
			createTargetTable (conn, dbType, normalizedTableName);

			String insertSql = buildInsertSql (normalizedTableName);
			log ("Prepared INSERT statement: " + insertSql);

			try (PreparedStatement ps = conn.prepareStatement (insertSql))
			{
				int fileIndex = 0;
				DataFormatter formatter = new DataFormatter (Locale.getDefault ());

				for (java.util.Map.Entry <File, java.util.List <SheetValidationResult>> entry : sheetsByFile
				        .entrySet ())
				{
					File file = entry.getKey ();
					java.util.List <SheetValidationResult> sheetResults = entry.getValue ();
					fileIndex++;

					int totalSheetsInFile = sheetResults.size ();
					int sheetIndex = 0;

					log ("Opening file: " + file.getName ());

					try (FileInputStream fis = new FileInputStream (file);
					     Workbook workbook = WorkbookFactory.create (fis))
					{

						for (SheetValidationResult svr : sheetResults)
						{
							sheetIndex++;
							String sheetName = svr.getSheetName ();
							if ("<all sheets>".equals (sheetName))
							{
								// Was an error placeholder in step 3
								continue;
							}

							Sheet sheet = workbook.getSheet (sheetName);
							if (sheet == null)
							{
								log ("  Skipping sheet '" + sheetName + "' (not found).");
								continue;
							}

							notifySheetStarted (fileIndex, totalFiles, sheetIndex, totalSheetsInFile, file, sheetName);

							importSheetData (sheet, formatter, ps);

							notifySheetCompleted (fileIndex, totalFiles, sheetIndex, totalSheetsInFile, file,
							                      sheetName);
							log ("  " + file.getName () + "@" + sheetName + " => completed.");
						}
					}
				}

				log ("Committing transaction...");
				conn.commit ();
				log ("Transaction committed.");

				updateImportedFileList ();
				log ("Updated traspasados_a_BBDD.txt.");
			}
			catch (Exception e)
			{
				log ("Error during import. Rolling back transaction...");
				conn.rollback ();
				log ("Transaction rolled back.");
				throw e;
			}
		}
	}

	private void validateConfiguration ()
	{
		if (config.getSelectedFiles () == null || config.getSelectedFiles ().isEmpty ())
		{
			throw new IllegalStateException ("No selected files in configuration.");
		}
		if (config.getDataStartRow () == null || config.getDataStartRow () <= 0)
		{
			throw new IllegalStateException ("Data start row not properly configured.");
		}
		if (config.getDbType () == null)
		{
			throw new IllegalStateException ("Database type not set.");
		}
		if (config.getDbHost () == null || config.getDbHost ().trim ().isEmpty ())
		{
			throw new IllegalStateException ("Database host not set.");
		}
		if (config.getDbPort () <= 0)
		{
			throw new IllegalStateException ("Database port not set or invalid.");
		}
		if (config.getDbName () == null || config.getDbName ().trim ().isEmpty ())
		{
			throw new IllegalStateException ("Database name not set.");
		}
		if (config.getDbUser () == null || config.getDbUser ().trim ().isEmpty ())
		{
			throw new IllegalStateException ("Database user not set.");
		}
	}

	private java.util.Map <File, java.util.List <SheetValidationResult>> groupSheetsByFile (java.util.List <SheetValidationResult> sheets)
	{
		java.util.Map <File, java.util.List <SheetValidationResult>> map = new LinkedHashMap <> ();
		for (SheetValidationResult svr : sheets)
		{
			File file = svr.getFile ();
			java.util.List <SheetValidationResult> list = map.computeIfAbsent (file, f -> new ArrayList <> ());
			list.add (svr);
		}
		return map;
	}

	private String buildJdbcUrl (DbType dbType, String host, int port, String dbName)
	{
		if (dbType == DbType.MARIADB)
		{
			return "jdbc:mariadb://" + host + ":" + port + "/" + dbName;
		}
		else
		{
			return "jdbc:postgresql://" + host + ":" + port + "/" + dbName;
		}
	}

	private boolean tableExists (Connection conn, String tableName) throws SQLException
	{
		DatabaseMetaData meta = conn.getMetaData ();
		try (ResultSet rs = meta.getTables (null, null, tableName, new String[] {"TABLE" }))
		{
			return rs.next ();
		}
	}

	private void createTargetTable (Connection conn, DbType dbType, String tableName) throws SQLException
	{
		java.util.List <ColumnDefinition> cols = config.getColumns ();
		boolean autoIncrement = config.isAutoIncrement ();

		StringBuilder sb = new StringBuilder ();
		sb.append ("CREATE TABLE ").append (tableName).append (" (");

		boolean first = true;

		if (autoIncrement)
		{
			if (dbType == DbType.MARIADB)
			{
				sb.append ("id BIGINT AUTO_INCREMENT PRIMARY KEY");
			}
			else
			{
				sb.append ("id BIGSERIAL PRIMARY KEY");
			}
			first = false;
		}

		for (ColumnDefinition col : cols)
		{
			if (!first)
			{
				sb.append (", ");
			}
			first = false;

			String colName = normalizeIdentifier (col.getName ());
			// For now, all columns as TEXT.
			sb.append (colName).append (" TEXT");
		}

		sb.append (")");

		String ddl = sb.toString ();
		log ("Executing DDL: " + ddl);

		try (Statement st = conn.createStatement ())
		{
			st.executeUpdate (ddl);
		}
	}

	private String buildInsertSql (String tableName)
	{
		java.util.List <ColumnDefinition> cols = config.getColumns ();

		StringBuilder sb = new StringBuilder ();
		sb.append ("INSERT INTO ").append (tableName).append (" (");

		boolean first = true;
		// We do NOT include the auto-increment ID column in INSERT
		for (ColumnDefinition col : cols)
		{
			if (!first)
			{
				sb.append (", ");
			}
			first = false;
			sb.append (normalizeIdentifier (col.getName ()));
		}

		sb.append (") VALUES (");

		first = true;
		for (int i = 0; i < cols.size (); i++)
		{
			if (!first)
			{
				sb.append (", ");
			}
			first = false;
			sb.append ("?");
		}
		sb.append (")");

		return sb.toString ();
	}

	private void importSheetData (Sheet sheet, DataFormatter formatter, PreparedStatement ps) throws SQLException
	{

		Integer headerRow = config.getHeaderRow (); // 1-based or null
		Integer dataStartRow = config.getDataStartRow (); // 1-based
		boolean fillEmpty = config.isFillEmptyCells ();

		int headerRowIndex = (headerRow != null && headerRow > 0)? (headerRow - 1) : -1;
		int dataStartIndex = (dataStartRow != null? dataStartRow - 1 : 0);

		int lastRow = sheet.getLastRowNum ();
		int columnCount = config.getColumns ().size ();

		String[] previousRowValues = new String[columnCount];

		for (int r = dataStartIndex; r <= lastRow; r++)
		{
			Row row = sheet.getRow (r);

			String[] currentValues = new String[columnCount];
			boolean rowHasAnyValue = false;

			for (int c = 0; c < columnCount; c++)
			{
				Cell cell = (row != null)? row.getCell (c, MissingCellPolicy.RETURN_BLANK_AS_NULL) : null;

				String value;
				if (cell == null)
				{
					value = "";
				}
				else
				{
					value = formatter.formatCellValue (cell);
				}

				if ((value == null || value.trim ().isEmpty ()) && fillEmpty)
				{
					if (previousRowValues[c] != null)
					{
						value = previousRowValues[c];
					}
				}

				if (value != null && !value.trim ().isEmpty ())
				{
					rowHasAnyValue = true;
				}

				currentValues[c] = value;
			}

			if (!rowHasAnyValue)
			{
				continue;
			}

			for (int c = 0; c < columnCount; c++)
			{
				String v = currentValues[c];
				if (v == null)
				{
					ps.setNull (c + 1, Types.VARCHAR);
				}
				else
				{
					ps.setString (c + 1, v);
				}
			}

			ps.addBatch ();
			previousRowValues = currentValues;
		}

		ps.executeBatch ();
	}

	private void updateImportedFileList ()
	{
		File dataDir = config.getDataDirectory ();
		if (dataDir == null)
		{
			log ("Data directory not defined; cannot update traspasados_a_BBDD.txt");
			return;
		}

		File txtFile = new File (dataDir, "traspasados_a_BBDD.txt");
		Set <String> existing = new LinkedHashSet <> ();

		if (txtFile.exists ())
		{
			try (BufferedReader br = new BufferedReader (new FileReader (txtFile)))
			{
				String line;
				while ((line = br.readLine ()) != null)
				{
					String trimmed = line.trim ();
					if (!trimmed.isEmpty ())
					{
						existing.add (trimmed);
					}
				}
			}
			catch (IOException e)
			{
				log ("Warning: could not read existing traspasados_a_BBDD.txt: " + e.getMessage ());
			}
		}

		for (File f : config.getSelectedFiles ())
		{
			existing.add (f.getName ());
		}

		try (PrintWriter pw = new PrintWriter (new FileWriter (txtFile, false)))
		{
			for (String name : existing)
			{
				pw.println (name);
			}
		}
		catch (IOException e)
		{
			log ("Warning: could not update traspasados_a_BBDD.txt: " + e.getMessage ());
		}
	}

	// === Helpers ===

	private String normalizeIdentifier (String raw)
	{
		if (raw == null)
		{
			return "col";
		}
		String s = raw.trim ().toLowerCase (Locale.ROOT);
		s = s.replaceAll ("[^a-z0-9_]", "_");
		if (s.isEmpty ())
		{
			s = "col";
		}
		return s;
	}

	private void log (String message)
	{
		if (listener != null)
		{
			listener.onLog (message + "\n");
		}
	}

	private void notifySheetStarted (int fileIndex, int totalFiles, int sheetIndex, int totalSheets, File file,
	                                 String sheetName)
	{
		if (listener != null)
		{
			listener.onSheetStarted (fileIndex, totalFiles, sheetIndex, totalSheets, file, sheetName);
		}
	}

	private void notifySheetCompleted (int fileIndex, int totalFiles, int sheetIndex, int totalSheets, File file,
	                                   String sheetName)
	{
		if (listener != null)
		{
			listener.onSheetCompleted (fileIndex, totalFiles, sheetIndex, totalSheets, file, sheetName);
		}
	}

	private void notifyCompleted ()
	{
		if (listener != null)
		{
			listener.onCompleted ();
		}
	}

	private void notifyError (Exception e)
	{
		if (listener != null)
		{
			listener.onError (e);
		}
	}
}
