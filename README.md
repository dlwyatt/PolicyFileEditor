# PolicyFileEditor
PowerShell functions and DSC resource wrappers around the TJX.PolFileEditor.PolFile .NET class.

This is for modifying registry.pol files (Administrative Templates) of local GPOs.  The .NET class code and examples of the original usage can be found at https://gallery.technet.microsoft.com/Read-or-modify-Registrypol-778fed6e .

It was written when I was still very new to both C# and PowerShell, and is pretty ugly / painful to use.  The new functions make this less of a problem, and the DSC resource wrapper around the functions will give us some capability to manage user-specific settings via DSC (something that's come up in discussions on a mailing list recently.)
