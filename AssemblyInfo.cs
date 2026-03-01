using System.Reflection;
using System.Runtime.CompilerServices;
using System;

//
// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
//
[assembly: AssemblyTitle("Created by Network International L.L.C. +233331400  statement@emp-group.com")]
[assembly: AssemblyDescription("Network International L.L.C.")]
[assembly: AssemblyConfiguration("Network International L.L.C.")]
[assembly: AssemblyCompany("Network International L.L.C.")]
[assembly: AssemblyProduct("Network International L.L.C.")]
[assembly: AssemblyCopyright("Network International L.L.C. +233331400  statement@emp-group.com")]
[assembly: AssemblyTrademark("Network International L.L.C.")]
[assembly: AssemblyCulture("")]

//[assembly: AssemblyTitle("Created by Emerging Markets Payments Africa (A network International Group Company) +233331400  statement@emp-group.com")]
//[assembly: AssemblyDescription("Emerging Markets Payments Africa (A network International Group Company)")]
//[assembly: AssemblyConfiguration("Emerging Markets Payments Africa (A network International Group Company)")]
//[assembly: AssemblyCompany("Emerging Markets Payments Africa (A network International Group Company)")]
//[assembly: AssemblyProduct("Emerging Markets Payments Africa (A network International Group Company)")]
//[assembly: AssemblyCopyright("Created by Emerging Markets Payments Africa (A network International Group Company) +233331400  statement@emp-group.com")]
//[assembly: AssemblyTrademark("Emerging Markets Payments Africa (A network International Group Company)")]
//[assembly: AssemblyCulture("")]

//[assembly: AssemblyTitle("Created by Network International +233331400  statement@emp-group.com")]
//[assembly: AssemblyDescription("Network International")]
//[assembly: AssemblyConfiguration("Network International")]
//[assembly: AssemblyCompany("Network International")]
//[assembly: AssemblyProduct("Network International")]
//[assembly: AssemblyCopyright("Created by Network International +233331400  statement@emp-group.com")]
//[assembly: AssemblyTrademark("Network International")]
//[assembly: AssemblyCulture("")]			

//
// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Revision and Build Numbers 
// by using the '*' as shown below:
//[assembly: AssemblyVersion("1.0.*")]

[assembly: AssemblyVersion("13.4")]

//
// In order to sign your assembly you must specify a key to use. Refer to the 
// Microsoft .NET Framework documentation for more information on assembly signing.
//
// Use the attributes below to control which key is used for signing. 
//
// Notes: 
//   (*) If no key is specified, the assembly is not signed.
//   (*) KeyName refers to a key that has been installed in the Crypto Service
//       Provider (CSP) on your machine. KeyFile refers to a file which contains
//       a key.
//   (*) If the KeyFile and the KeyName values are both specified, the 
//       following processing occurs:
//       (1) If the KeyName can be found in the CSP, that key is used.
//       (2) If the KeyName does not exist and the KeyFile does exist, the key 
//           in the KeyFile is installed into the CSP and used.
//   (*) In order to create a KeyFile, you can use the sn.exe (Strong Name) utility.
//       When specifying the KeyFile, the location of the KeyFile should be
//       relative to the project output directory which is
//       %Project Directory%\obj\<configuration>. For example, if your KeyFile is
//       located in the project directory, you would specify the AssemblyKeyFile 
//       attribute as [assembly: AssemblyKeyFile("..\\..\\mykey.snk")]
//   (*) Delay Signing is an advanced option - see the Microsoft .NET Framework
//       documentation for more information on this.
//
[assembly: AssemblyDelaySign(false)]
[assembly: AssemblyKeyFile("")]
[assembly: AssemblyKeyName("")]



#region " Helper class to get information for the About form. "

// This class uses the System.Reflection.Assembly class to
// access assembly meta-data
// This class is ! a normal feature of AssemblyInfo.cs

public class AssemblyInfo 
{
	// Used by Helper Functions to access information from Assembly Attributes

	private Type myType;

	public AssemblyInfo() 
	{
		myType = typeof(frmBasLogin);
	}

	public string AsmName 
	{
		get 
		{
			return myType.Assembly.GetName().Name.ToString();
		}
	}

	public string AsmFQName 
	{
		get 
		{
			return myType.Assembly.GetName().FullName.ToString();
		}
	}

	public string CodeBase 
	{
		get 
		{
			return myType.Assembly.CodeBase;
		}
	}

	public string Copyright 
	{
		get 
		{
			Type at = typeof(AssemblyCopyrightAttribute);
			object[] r = myType.Assembly.GetCustomAttributes(at, false);
			AssemblyCopyrightAttribute ct = (AssemblyCopyrightAttribute) r[0];
			return ct.Copyright;
		}
	}

	public string Company 
	{
		get 
		{
			Type at = typeof(AssemblyCompanyAttribute);
			object[] r = myType.Assembly.GetCustomAttributes(at, false);
			AssemblyCompanyAttribute ct = (AssemblyCompanyAttribute) r[0];
			return ct.Company;
		}
	}

	public string Description 
	{
		get 
		{
			Type at = typeof(AssemblyDescriptionAttribute);
			object[] r = myType.Assembly.GetCustomAttributes(at, false);
			AssemblyDescriptionAttribute da = (AssemblyDescriptionAttribute) r[0];
			return da.Description;
		}
	}

	public string Product 
	{
		get 
		{
			Type at = typeof(AssemblyProductAttribute);
			object[] r = myType.Assembly.GetCustomAttributes(at, false);
			AssemblyProductAttribute pt = (AssemblyProductAttribute) r[0];
			return pt.Product;
		}
	}

	public string Title 
	{
		get 
		{
			Type at = typeof(AssemblyTitleAttribute);
			object[] r = myType.Assembly.GetCustomAttributes(at, false);
			AssemblyTitleAttribute ta = (AssemblyTitleAttribute) r[0];
			return ta.Title;
		}
	}

	public string Version 
	{
		get 
		{
			return myType.Assembly.GetName().Version.ToString();
		}
	}

}

#endregion


/*
AssemblyInfo ainfo = new AssemblyInfo();
this.lblTitle.Text = ainfo.Title;
this.lblVersion.Text = string.Format("Version {0}", ainfo.Version);
this.lblCopyright.Text = ainfo.Copyright;
this.lblDescription.Text = ainfo.Description;
this.lblCodebase.Text = ainfo.CodeBase;
 */
