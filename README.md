# TransferSheetGenerator
Generates transfer sheets for use on the Labcyte Echo acoustic liquid handler using a collection of chemicals as defined by a database/spreadsheet.

The Labcyte Echo uses a row/column notation to define positions for microplate wells. The Echo is a liquid handler that uses acoustics to transfer discrete volumes of fluid and takes user inputs in the form of a spreadsheet containing desired source and destination plate/well locations. Using a spreadsheet or database to define the locations of chemicals stored in microplates (and their associated metadata) in a high-throughput setting, this code gives users control over transferring chemicals of a specific class for high-throughput experiments (regardless of the number of chemicals in any given class).

Chemical Transfer Tab
The easy-to-use interface allows users to select which class of chemicals they desire by entering in a discrete volume in nanoliters. It also gives the option to omit empty/unused wells in a plate to keep transfer sheets succinct.

Assay Transfer Tab
Destination plate requirements are different for high-throughput assay experiments, thus, an assay transfer tab was created to allow users to make transfer sheets to transfer compounds from the chemical collection into their assays. Again, users can select which class of chemicals they'd like to screen in their assay, and can define a volume in nanoliters. Empty/unused wells in plates of the chemical collection can be omitted.
