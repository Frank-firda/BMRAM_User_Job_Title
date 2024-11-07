# BMRAM_User_Job_Title

### Overview
**BMRAM_User_Job_Title** is a tool that searches for names in an Excel sheet and retrieves their corresponding email addresses and job titles from the Outlook Global Address List (GAL). The app writes this information into specified columns in the Excel sheet or generates a new one based on your preferences.

### Features
- Retrieves and writes emails and job titles from Outlook’s GAL.
- Customizable output columns and sheet name.
- Optional settings for writing to a new Excel file or overwriting the existing one.
- Automatic launch of the Excel sheet upon completion.
- Halt on error option with customizable user actions.
- Configurable settings via a separate settings executable.

### Installation
1. Download and install the executable files.
2. Uninstall via **Programs and Features** in your system settings if needed.

### How to Use

#### Main App (`BMRAM User, Job Title.exe`)
1. Close the Excel file you want to read from.
2. Run `BMRAM User, Job Title.exe`.
3. Select the Excel file (.xlsx) in the file dialog.
4. The app will begin processing, displaying a progress bar.
5. If a person in the list isn’t found, "Not found" will be recorded in the Excel sheet for both email and job title.

#### Settings (`Settings.exe`)
1. Run `Settings.exe` to configure the app:
   - **Email and Job Title Writing**: Choose whether to include emails and/or job titles.
   - **Column Filter**: Set a filter to apply to the columns.
   - **Write Mode**: Select whether to overwrite the current Excel sheet or save to a new one.
   - **Automatic Excel Launch**: Choose if the Excel file should open upon completion.
   - **Column Assignments**: Customize the columns used for each item:
     - **Email Column**: Default `B`
     - **Job Title Column**: Default `C`
     - **Name Column**: Default `A`
     - **Sheet Name**: Specify the sheet to read and write to (same sheet required for both).
   - **Halt on Error**: If enabled, the program will notify you if an error occurs, allowing the following responses:
     - `R` - Retry (may not resolve the issue)
     - `S` - Skip and continue to the next person
     - `C` - Cancel the operation and close the program
   - **Save Settings**: Press “Save” to apply changes.

### Troubleshooting
If `Halt on Error` is off, the app auto-skips any entries without a found email or job title. When `Halt on Error` is on, follow the prompt actions as described above.

#### Uninstall
To uninstall, go to **Programs and Features** and select **Uninstall**.

### Known Issues
- Occasionally, the first name in the list may not retrieve any data, leaving email and job title cells blank.



