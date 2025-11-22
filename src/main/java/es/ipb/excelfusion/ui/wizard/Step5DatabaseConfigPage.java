
package es.ipb.excelfusion.ui.wizard;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

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
import org.eclipse.swt.widgets.Text;


/**
 * Step 5 of the wizard:
 * - Database configuration (MariaDB / PostgreSQL)
 * - Host, port, database name, user, password
 * - "Create DB if not existing" option
 * - "Test connection" button
 */
public class Step5DatabaseConfigPage implements WizardPage
{

	public enum DbType
	{
		MARIADB, POSTGRESQL
	}

	private Composite control;

	private Combo	  dbTypeCombo;
	private Text	  hostText;
	private Text	  portText;
	private Text	  dbNameText;
	private Text	  userText;
	private Text	  passwordText;
	private Button	  createDbIfMissingCheckbox;
	private Button	  testConnectionButton;
	private Label	  testResultLabel;

	// Cached values (optional, in case you want them later)
	private DbType	  selectedDbType = DbType.MARIADB;

	public Step5DatabaseConfigPage ()
	{
	}

	@Override
	public String getTitle ()
	{
		return "Database configuration";
	}

	@Override
	public String getDescription ()
	{
		return "Configure the target database (MariaDB or PostgreSQL) and test the connection.";
	}

	@Override
	public void createControl (Composite parent)
	{
		control = new Composite (parent, SWT.NONE);
		control.setLayout (new GridLayout (1, false));

		createDbConfigGroup (control);
		createBottomSection (control);
	}

	private void createDbConfigGroup (Composite parent)
	{
		Group group = new Group (parent, SWT.NONE);
		group.setText ("Connection settings");
		group.setLayoutData (new GridData (SWT.FILL, SWT.TOP, true, false));
		group.setLayout (new GridLayout (2, false));

		// DB type
		Label dbTypeLabel = new Label (group, SWT.NONE);
		dbTypeLabel.setText ("Database type:");

		dbTypeCombo = new Combo (group, SWT.DROP_DOWN | SWT.READ_ONLY);
		dbTypeCombo.setItems (new String[] {"MariaDB", "PostgreSQL" });
		dbTypeCombo.select (0); // default MariaDB
		dbTypeCombo.addSelectionListener (new SelectionAdapter ()
		{
			@Override
			public void widgetSelected (SelectionEvent e)
			{
				selectedDbType = (dbTypeCombo.getSelectionIndex () == 0)? DbType.MARIADB : DbType.POSTGRESQL;
				adjustDefaultPortIfEmpty ();
			}
		});

		// Host
		Label hostLabel = new Label (group, SWT.NONE);
		hostLabel.setText ("Host:");

		hostText = new Text (group, SWT.BORDER);
		hostText.setLayoutData (new GridData (SWT.FILL, SWT.CENTER, true, false));
		hostText.setText ("localhost");

		// Port
		Label portLabel = new Label (group, SWT.NONE);
		portLabel.setText ("Port:");

		portText = new Text (group, SWT.BORDER);
		portText.setLayoutData (new GridData (SWT.FILL, SWT.CENTER, true, false));
		portText.setText ("3306"); // default for MariaDB

		// Database name
		Label dbNameLabel = new Label (group, SWT.NONE);
		dbNameLabel.setText ("Database name:");

		dbNameText = new Text (group, SWT.BORDER);
		dbNameText.setLayoutData (new GridData (SWT.FILL, SWT.CENTER, true, false));

		// User
		Label userLabel = new Label (group, SWT.NONE);
		userLabel.setText ("User:");

		userText = new Text (group, SWT.BORDER);
		userText.setLayoutData (new GridData (SWT.FILL, SWT.CENTER, true, false));

		// Password
		Label passwordLabel = new Label (group, SWT.NONE);
		passwordLabel.setText ("Password:");

		passwordText = new Text (group, SWT.BORDER | SWT.PASSWORD);
		passwordText.setLayoutData (new GridData (SWT.FILL, SWT.CENTER, true, false));

		// Create DB if missing
		Label createDbLabel = new Label (group, SWT.NONE);
		createDbLabel.setText ("Create DB if not existing:");

		createDbIfMissingCheckbox = new Button (group, SWT.CHECK);
		createDbIfMissingCheckbox.setSelection (false);
	}

	private void createBottomSection (Composite parent)
	{
		Composite bottom = new Composite (parent, SWT.NONE);
		bottom.setLayoutData (new GridData (SWT.FILL, SWT.TOP, true, false));
		bottom.setLayout (new GridLayout (2, false));

		testConnectionButton = new Button (bottom, SWT.PUSH);
		testConnectionButton.setText ("Test connection");
		testConnectionButton.setLayoutData (new GridData (SWT.LEFT, SWT.CENTER, false, false));
		testConnectionButton.addSelectionListener (new SelectionAdapter ()
		{
			@Override
			public void widgetSelected (SelectionEvent e)
			{
				testConnection ();
			}
		});

		testResultLabel = new Label (bottom, SWT.WRAP);
		GridData gd = new GridData (SWT.FILL, SWT.CENTER, true, false);
		gd.widthHint = 400;
		testResultLabel.setLayoutData (gd);
		testResultLabel.setText ("");
	}

	@Override
	public Composite getControl ()
	{
		return control;
	}

	@Override
	public void onEnter ()
	{
		// In a later iteration we could load config from a per-directory file here.
	}

	@Override
	public boolean onLeave ()
	{
		// Validate configuration before moving on
		if (!validateConfig (false))
		{
			return false;
		}
		return true;
	}

	@Override
	public boolean canGoNext ()
	{
		// Allow Next, but real validation happens in onLeave()
		return true;
	}

	@Override
	public boolean canGoBack ()
	{
		return true;
	}

	private void adjustDefaultPortIfEmpty ()
	{
		String portStr = portText.getText ().trim ();
		if (!portStr.isEmpty ())
		{
			return;
		}
		if (selectedDbType == DbType.MARIADB)
		{
			portText.setText ("3306");
		}
		else
		{
			portText.setText ("5432");
		}
	}

	private boolean validateConfig (boolean showDialogs)
	{
		String host = hostText.getText ().trim ();
		String portStr = portText.getText ().trim ();
		String dbName = dbNameText.getText ().trim ();
		String user = userText.getText ().trim ();
		// password can be empty

		if (host.isEmpty ())
		{
			if (showDialogs)
			{
				showError ("Invalid host", "Host is required.");
			}
			return false;
		}
		if (portStr.isEmpty ())
		{
			if (showDialogs)
			{
				showError ("Invalid port", "Port is required.");
			}
			return false;
		}
		try
		{
			int port = Integer.parseInt (portStr);
			if (port <= 0 || port > 65535)
			{
				if (showDialogs)
				{
					showError ("Invalid port", "Port must be a number between 1 and 65535.");
				}
				return false;
			}
		}
		catch (NumberFormatException e)
		{
			if (showDialogs)
			{
				showError ("Invalid port", "Port must be a valid integer.");
			}
			return false;
		}

		if (dbName.isEmpty ())
		{
			if (showDialogs)
			{
				showError ("Invalid database name", "Database name is required.");
			}
			return false;
		}

		if (user.isEmpty ())
		{
			if (showDialogs)
			{
				showError ("Invalid user", "User is required.");
			}
			return false;
		}

		return true;
	}

	private void testConnection ()
	{
		if (!validateConfig (true))
		{
			return;
		}

		DbType dbType = getDbType ();
		String host = hostText.getText ().trim ();
		String port = portText.getText ().trim ();
		String dbName = dbNameText.getText ().trim ();
		String user = userText.getText ().trim ();
		String password = passwordText.getText ();

		boolean createDb = createDbIfMissingCheckbox.getSelection ();

		testResultLabel.setText ("Testing connection...");
		testResultLabel.getParent ().layout ();

		String jdbcUrl = buildJdbcUrl (dbType, host, port, dbName);

		try
		{
			loadDriverClass (dbType);

			// Try connecting directly to the DB
			try (Connection conn = DriverManager.getConnection (jdbcUrl, user, password))
			{
				testResultLabel.setText ("Connection successful.");
				showInfo ("Connection test", "Successfully connected to the database.");
				return;
			}
			catch (SQLException ex)
			{
				// If connection fails and createDb is enabled, we may attempt to create DB
				if (createDb)
				{
					boolean created = tryCreateDatabase (dbType, host, port, dbName, user, password);
					if (created)
					{
						testResultLabel.setText ("Database created and connection successful.");
						showInfo ("Connection test",
						          "Database did not exist but was created successfully, and the connection works.");
						return;
					}
				}
				throw ex;
			}

		}
		catch (SQLException e)
		{
			e.printStackTrace ();
			testResultLabel.setText ("Connection failed: " + e.getMessage ());
			showError ("Connection failed", "Could not connect to database:\n\n" + e.getMessage ());
		}
		catch (ClassNotFoundException e)
		{
			e.printStackTrace ();
			testResultLabel.setText ("Driver not found: " + e.getMessage ());
			showError ("Driver not found", "JDBC driver not found for " + dbType + ":\n\n" + e.getMessage ());
		}
	}

	private void loadDriverClass (DbType dbType) throws ClassNotFoundException
	{
		if (dbType == DbType.MARIADB)
		{
			Class.forName ("org.mariadb.jdbc.Driver");
		}
		else
		{
			Class.forName ("org.postgresql.Driver");
		}
	}

	private String buildJdbcUrl (DbType dbType, String host, String port, String dbName)
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

	/**
	 * Very simple attempt to create the database if not existing.
	 * Implementation is intentionally minimal; in a real system you might want
	 * a specific "admin" database or configuration.
	 */
	private boolean tryCreateDatabase (DbType dbType, String host, String port, String dbName, String user,
	                                   String password)
	{

		String adminDbName = (dbType == DbType.MARIADB)? "mysql" : "postgres";
		String adminUrl = buildJdbcUrl (dbType, host, port, adminDbName);

		String createStatement = "CREATE DATABASE " + dbName;

		try (Connection adminConn = DriverManager.getConnection (adminUrl, user, password);
		     java.sql.Statement stmt = adminConn.createStatement ())
		{

			stmt.executeUpdate (createStatement);

			// Now test connection to the newly created DB:
			String newDbUrl = buildJdbcUrl (dbType, host, port, dbName);
			try (Connection conn = DriverManager.getConnection (newDbUrl, user, password))
			{
				return true;
			}

		}
		catch (SQLException e)
		{
			e.printStackTrace ();
			showError ("Create DB failed", "Failed to create database '" + dbName + "':\n\n" + e.getMessage ());
			return false;
		}
	}

	private void showError (String title, String message)
	{
		MessageBox mb = new MessageBox (control.getShell (), SWT.ICON_ERROR | SWT.OK);
		mb.setText (title);
		mb.setMessage (message);
		mb.open ();
	}

	private void showInfo (String title, String message)
	{
		MessageBox mb = new MessageBox (control.getShell (), SWT.ICON_INFORMATION | SWT.OK);
		mb.setText (title);
		mb.setMessage (message);
		mb.open ();
	}

	// === Public getters for Step 6 ===

	public DbType getDbType ()
	{
		return (dbTypeCombo.getSelectionIndex () == 0)? DbType.MARIADB : DbType.POSTGRESQL;
	}

	public String getHost ()
	{
		return hostText.getText ().trim ();
	}

	public int getPort ()
	{
		try
		{
			return Integer.parseInt (portText.getText ().trim ());
		}
		catch (NumberFormatException e)
		{
			return -1;
		}
	}

	public String getDatabaseName ()
	{
		return dbNameText.getText ().trim ();
	}

	public String getUser ()
	{
		return userText.getText ().trim ();
	}

	public String getPassword ()
	{
		return passwordText.getText ();
	}

	public boolean isCreateDbIfMissing ()
	{
		return createDbIfMissingCheckbox.getSelection ();
	}
}
