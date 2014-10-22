Stoichiometry-Calculator
========================

Component-Based Programming .NET/ VBA Macros

Create a library assembly using C# that provides stoichiometric services and a windows client to demonstrate the library.

This Project:
- develop a .NET class library assembly
- use the interface type in C# to define an immutable interface
- generate a strong-named assembly and share it in the GAC
- use the .NET interop layer to combine managed and unmanaged components
- use ADO .NET to retrieve and modify data within a database

How I Signed the Stoichiometry Assembly
- Open the Stoichiometry project in Visual Studio.
- Add a reference to the key file in the project’s properties.
- Right-click on the project in the Solution Explorer and select Properties
- Select the Signing page
- Click on Sign the Assembly
- Select <Browse…> under Choose a Strong Name Key File, or if you haven’t already created a key file select <New…> to create one.
- Type in a password if you choose to password-protect your key file.


How I Generated and registered my CCW
You can have Visual Studio automatically generate your COM Callable Wrapper (CCW) and register it when you build your Stoichiometry assembly as follows:
- Open the Stoichiometry project in Visual Studio.
- In the Property pages for the Stoichiometry assembly:
- select the Application page, click on Assembly Information and click on Make assembly COM-Visible
- select the Build page and click on Register for COM interop
- Rebuild the Stoichiometry assembly – this will generate and register the CCW (stoichiometry.tlb). Do not move this file! Visual Studio makes an entry into the System Registry in Windows which incorporates the path to the file.

Instructions for making the excel client refer to my Stoichiometry assembly
*Note: Before opening the client document in Excel the Stoichiometry.mdb
database must installed be in the working directory for Excel (probably
‘Documents’)
Using Microsoft Excel:
- Open the Excel client
- If you don’t see a menu tab called Developer, click on File in the menu
ribbon then select Options. On the Custom Ribbon page check Developer
on the right side and click OK.
- On the Developer menu, click on the Visual Basic tool.
- In the Visual Basic Editor, select Tools->References…
- In the References – VBAProject dialog find and checkmark Stoichiometry
in the Available References list and click OK.
- Close the Visual Basic Editor window.
- Above the formula bar you should see a line that says Security Warning
as follows:
Click the Enable Content button to the right of this.
- If everything is working correctly the Weight and Normalized Formula
columns on the Your Calculator page should show valid results.
