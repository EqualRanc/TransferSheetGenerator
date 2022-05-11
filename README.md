# TransferSheetGenerator
Generates transfer sheets for use on the Labcyte Echo acoustic liquid handler using a collection of chemicals as defined by a database/spreadsheet.

The Labcyte Echo uses a row/column notation to define positions for microplate wells. The Echo is a liquid handler that uses acoustics to transfer discrete volumes of fluid and takes user inputs in the form of a spreadsheet containing desired source and destination plate/well locations. Using a spreadsheet or database to define the locations of chemicals stored in microplates (and their associated metadata) in a high-throughput setting, this code gives users control over transferring chemicals of a specific class for high-throughput experiments (regardless of the number of chemicals in any given class).

Chemical Transfer Tab
The easy-to-use interface allows users to select which class of chemicals they desire by entering in a discrete volume in nanoliters. It also gives the option to omit empty/unused wells in a plate to keep transfer sheets succinct.

Assay Transfer Tab
Destination plate requirements are different for high-throughput assay experiments, thus, an assay transfer tab was created to allow users to make transfer sheets to transfer compounds from the chemical collection into their assays. Again, users can select which class of chemicals they'd like to screen in their assay, and can define a volume in nanoliters. Empty/unused wells in plates of the chemical collection can be omitted.

# Instructions
1. Run the TransferSheetGenerator.py file.
2. Select either the Chemical Transfer tab or the Assay Transfer Tab depending on experiment desired.
3. Browse to local disk location of the chemical database file (ex. 20220419_Chemical_Database_Example.xlsx)
4. Enter desired transfer volumes in nanoliters into each of the appropriate boxes for the desired chemical class(es)
5. Enter name of the experiment in the "Enter metadata..." field
6. Check the "Remove empty rows?" option for a more concise transfer sheet (and smaller file size)
7. Click "Generate Chemical Transfer Sheet" or "Generate Assay Transfer Sheet"
8. When complete, the status window will indicate the transfer sheet generation was successful and indicate location.
    The status window will also print an error message if it was unsuccessful in the event the tool needs to be modified to accommodate new classes of chemicals, new storage plates, etc.
9. Click "Exit (Chemical Tab)" or "Exit (Assay Tab)" to close.
