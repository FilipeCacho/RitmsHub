# How to Install

To install the code, all you need to do is clone the repository to Visual Studio and run the project. Everything should be in place, and the necessary NuGet packages should be installed automatically.

# About Table Locks

**Option 1** allows the user to create and update Business Units (BUs) and Teams. Normally, this process can take about 5 seconds per BU or Team. During this time, Dynamics might impose a table lock on the BU or Team table to ensure data integrity, meaning other users may experience delays when accessing these tables. During normal project functionality, all locks are released automatically by Dynamics when the process finishes. However, if an unexpected error occurs or if the console is killed while the process is ongoing, a table lock might persist.

**Option 12** in the console is a best-effort attempt to release these locks but is not guaranteed to work. No one knows how to manually release these locks, as they are managed by Dynamics, and only Microsoft can presumably lift them. In the past, such locks have lifted themselves after 30 minutes to 1 hour.

To summarize, the console is safe to use and won’t cause significant disruption if used to create a few BUs and Teams (10-20 max). However, users might experience system unresponsiveness if more than that are created. If you decide to create dozens or more, do not stop the creation process midway as this will cause a table lock. It’s best to execute large operations after working hours, preferably no sooner than 10 PM.

## Important Notes

This program revolves around a file called **dataCenter.xlsx**. This file has 3 worksheets:

- **Login**: Contains your connection settings. You must change the **Username and Password** columns for each environment you want to access. **Do not modify any other columns, as they contain necessary info to connect to Microsoft CRM.**
- **Create Teams**: This worksheet is where you input the information for the teams you want to create. Upon opening the sheet for the first time, rows 1 and 2 contain placeholder information, which should be removed. These placeholders simply show how to structure your data.
  - **Columns A-D** must be filled out with the required information, or else the team cannot be created. **Column E** is used to run flows for Work Orders and Notifications, but it can be filled with any value if you don’t plan on running these flows.
- **Assign Teams**: This worksheet allows you to assign existing teams to active users. Specify the username and the team to be assigned. For example, if you need to give a user access to a team restricted by a contractor, you must assign both the **master data team (equipo dados maestros)** and the **Contrata** team. The information in this sheet is placeholder and should be replaced.

### Minor Details

Since this project runs on a corporate PC, certain restrictions apply. For instance, resizing the console is limited. At the top of the console, the first line tells you which environment you're connected to (PRD, DEV, or PRE), but you may need to scroll up to see it. By default, the project starts connected to the PRD environment.

My initial plan was to compile the code into a single executable, but this presents two problems:
1. A compiled executable triggers Windows protections, and though the code runs, it doesn't work because it’s not signed (signing costs money).
2. Running the code as an administrator is another option, but most of us can’t escalate privileges without contacting the service desk. So, the simplest solution is to distribute the project instead.

## How to Use the Project

After running the project, a menu with several options is presented:

### Create Teams Flow

- **Option 1**: Reads the information from the **Create Teams** worksheet, checks if it's properly formatted and complete, and displays the BU and teams that can be created. If the user confirms, the creation process begins:
  - First, it checks if the BU exists. If it does, it checks if the BU information matches what was entered in the worksheet. If any data differs, the BU is updated with the new information. **It is the user’s responsibility to ensure the information is accurate when updating an existing BU**.
  - If the BU doesn’t exist, it's created with the details from the worksheet.
  - The system then creates the team restricted by the indicated contractor. If the team already exists, any non-matching information is updated.
  - Finally, a proprietary team is created and associated with the respective BU. Again, if the team exists, it is updated based on the worksheet information.
  
- **Option 2**: Retrieves all active users from the teams listed in the **Create Teams** worksheet. It removes duplicates and displays only currently enabled users in the console, also generating Excel files for each BU in the **Generated Excels** folder in your Downloads. This list is retained in memory for subsequent operations.
  
- **Option 3**: (Requires **Option 2** to run first) After extracting users from each BU, it assigns the newly created contractor-restricted team to the respective users.

- **Option 4**: Uses the **contractor** column in **Create Teams** worksheet to **create views for Work Orders and Notifications** to run flows in XRM Toolbox. **The views are restricted to the specified contractor**. After creating the query, you can name it and verify the view in Dynamics or run the workflow in XRM Toolbox.

- **Option 5**: Ensures that users have access restricted by the correct contractor. If their current BU isn’t restricted, this option updates it accordingly and ensures their access is properly limited.

  ### Other flows

- **Option 6**: Assigns minimum roles and teams based on whether the user is in the EU or NA. If the user is an internal EU user, for example, they are assigned **Equipo conocimiento personal Interno** and the **RESCO Archive Read** role. It asks if you want to assign RESCO, and if so, the role and team are assigned. However, **this doesn’t update the user in Woodford**, which must be done manually.

- **Option 7**: A development-only feature that displays a list of workflows a user can run (not for normal use).

- **Option 8**: Copies BUs, Teams, and/or roles from one user to another. It’s a blind copy, so you don’t get to select specific permissions. If you replace the BU, all roles are lost unless copied over as well.

- **Option 9**: Stores a user’s roles in memory while you change their BU via the web interface. When the BU is updated, their saved roles can be reapplied.

- **Option 10**: Assigns teams from the **Assign Teams** worksheet. It's a blind assign—teams already assigned to a user are skipped.

- **Option 11**: Displays a user’s BU, teams, and roles. You can compare this with another user’s information, with common permissions highlighted.

- **Option 12**: Attempts to release table locks caused by an unexpected error when running **Option 1**. This is a best-effort approach with no guaranteed results. It should only be used in emergencies when BU-related errors occur.

- **Option 13**: Switches between environments by fetching your corresponding login details from the **Login** worksheet.

- **Option 0**: Safely closes the console at any time, as long as no information is being fetched or uploaded to Dynamics.
