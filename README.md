XPS WATERMARKING WITH DEVIARE
=============================
Software requeriments:

* Windows Vista SP2 with Platform Update or higher.
* Visual Studio 2008 with Service Pack 1
* MS Windows Platform SDK 7.1 or higher.
* .NET Framework 2.0
* Internet Explorer 9.0 or higher.

(!) Be sure to run the correct build for your platform. Use X86 build for 32-bit Windows systems,
and X64 for 64-bit Windows systems.


To build under VisualStudio 2008
--------------------------------

To build the code, open the IEPrintWatermark solution and select "Batch Build..." in the Build Menu. 
This is necessary for the build process to generate the binaries in the correct order of dependencies.

Run IEPrintWatermark. The default IE installation will be launched. Navigate to your desired page,
Print the page to an XPS-compatible printer; a sample watermark consisting of a light blue box and
a semitransparent text string will be superimposed on the page.


Notes on Source Code
--------------------

- The solution consists for two C# projects: IEPrintWatermark which initiates the hooking process
  and loads the custom C++ DLL and C# plugin, and IEPrintWatermarkHelperCS which is the plugin itself,
  doing all the internal XPS watermark drawing operations. The remaining IEPrintWatermarkHelper project code
  is C++ and it's just a routine to get a pointer to an internal interface.

- The Microsoft XPS PrintAPI is COM-based and it was imported using the .NET framework tools, from the
  SDK xpsobjectmodel.idl file. The converted file is MSXPS.DLL and it's found on the binary folders.

- Used XPS Document API is documented by Microsoft on: 
  http://msdn.microsoft.com/en-us/library/windows/desktop/dd316976(v=vs.85).aspx

- Additional code snippets were included to make easier the use of XPS Document API from C#. See for
  example "MakeXPSColor" function to generate a proper RGBA color value; and MakeMatrixTransform 
  for generating transformation matrices for rotating, translating and scaling the visual objects in
  the XPS document.

- For XPS Document, all font resources are embedded from file streams. However, a small font name
  to TTF filename mapping code is included (see GetSystemFontFileName function) for easier handling.
  Font name is case sensitive.

- XPS font size is specified in "ems", not pixels or points. 1 document inch equals to 96 XPS units.

- All object transformations are relative to page origin (top-left is 0,0).


