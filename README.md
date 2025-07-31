# VBADocMover
Document Copier using VBA

This VBA code is a file management tool designed to simplify and automate the process of copying client statements to their correct folders based on user-defined criteria. 

The macro reads a list of file names from an Excel sheet, verifies that the corresponding destination folders and account details are accurate, and copies the files from a designated source location to the specified destinations. It identifies and resolves broken folder paths by prompting users to correct them, updates file names dynamically by replacing substrings (e.g., updating the statement date), and ensures files are organized properly. Additionally, it tracks files successfully copied and lists those skipped due to duplication in the destination folder.

This code significantly improves efficiency, reducing a process that previously took 2–3 minutes per client to about 5 minutes for all files combined. It minimizes the risk of errors from manual file handling and offers flexibility by allowing users to add new file destinations and account details for unmapped files. Clear prompts and input boxes guide users through decisions like selecting folders or entering account names, ensuring the process adapts to a variety of scenarios. Furthermore, the macro provides detailed feedback on copied and skipped files, enabling users to address issues efficiently.
This tool is especially valuable for managing large-scale file transfers, such as organizing client statements, financial reports, or similar documents, where files must be categorized by account, date, or folder.
