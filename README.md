# Splitter

# Part One – Splitter
The Invoice Splitter makes it easier to manage your invoices without having to rename each file manually. The invoice is displayed, and you only need to change the invoice number. Once you're done, the files are automatically saved in the folder you choose.

# Part Two – Generator
The Invoice Generator helps you create multiple invoices much faster. You simply upload the invoice, choose how many invoices you want to generate, and select the format you need — Excel, PDF, or merged PDFs.

Both tools are designed to make the process easier, save time, and reduce the amount of manual work.

Pamja: 

<img width="1083" height="727" alt="1" src="https://github.com/user-attachments/assets/97a1eed8-bc2d-423d-b36b-cb30a1254db4" />


<img width="1083" height="742" alt="2" src="https://github.com/user-attachments/assets/af087453-9de5-422f-a23f-fa93fd58c38b" />


# Tech Stack
This project was developed using standard web technologies — HTML, CSS, and JavaScript — integrated through the Electron framework.

User Interface (Frontend)
HTML5 & CSS3 (Vanilla): Used to build the application's interface, including its structure, forms, buttons, and overall window design.
Vanilla JavaScript: Used to handle the application's logic and user interactions, without relying on frameworks such as React, Vue, or Angular.

Backend / Operating System:
Built on Node.js integrated with Electron to provide native desktop capabilities:
File System Access: Reads and updates Excel template files dynamically using exceljs.
PDF Processing: Merges generated invoices into a single PDF via pdf-lib.
OS Integration: Executes background PowerShell scripts to automate Microsoft Excel PDF export and handle file paths.

This approach combines the simplicity of standard web technologies with the ability to perform direct system-level operations through Electron and Node.js.
